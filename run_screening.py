"""
FOSSEE Summer Fellowship 2026 - DWSIM Python Automation Screening Task
Task 1: Python Automation of DWSIM

Author: Atharva Ramawat
Description:
    This script uses the DWSIM Automation API (via pythonnet / DWSIM's COM interface)
    to programmatically:
        Part A - Simulate a PFR for isomerization of n-pentane to isopentane
        Part B - Simulate a Distillation Column separating n-pentane / isopentane
        Part C - Perform parametric sweeps over key variables

    All results are written to results.csv. No GUI is launched.
"""

import sys
import os
import csv
import traceback
import itertools
import logging
from datetime import datetime

# ── Safe Path Setup ──────────────────────────────────────────────────────────
# Guarantees outputs are saved next to the script, even after changing directories
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ── Optional: matplotlib for plots (won't crash if unavailable) ─────────────
try:
    import matplotlib
    matplotlib.use("Agg")          # headless backend – no display needed
    import matplotlib.pyplot as plt
    PLOT_AVAILABLE = True
except ImportError:
    PLOT_AVAILABLE = False

# ── Logging setup ────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(os.path.join(SCRIPT_DIR, "run_screening.log"), mode="w"),
    ],
)
log = logging.getLogger(__name__)

# ─────────────────────────────────────────────────────────────────────────────
# 1.  DWSIM AUTOMATION BOOTSTRAP
# ─────────────────────────────────────────────────────────────────────────────

def find_dwsim_path():
    """Return the path to the DWSIM installation directory."""
    candidates = [
        r"C:\Users\Public\DWSIM8",
        r"C:\Program Files\DWSIM8",
        r"C:\Program Files (x86)\DWSIM8",
        r"C:\Users\Public\DWSIM7",
        r"C:\Program Files\DWSIM7",
        r"C:\Users\athar\AppData\Local\DWSIM", # Confirmed path
        os.environ.get("DWSIM_HOME", ""),
        "/opt/dwsim8",
        "/usr/local/dwsim8",
        os.path.expanduser("~/dwsim8"),
    ]
    for p in candidates:
        if p and os.path.isdir(p):
            log.info(f"DWSIM found at: {p}")
            return p
    raise EnvironmentError(
        "DWSIM installation not found. Set the DWSIM_HOME environment variable "
        "or install DWSIM to a standard location."
    )


def load_dwsim(dwsim_path: str):
    """
    Bootstrap pythonnet and load DWSIM assemblies.
    Returns the automation object used to create flowsheets.
    """
    try:
        import clr  
    except ImportError:
        raise ImportError(
            "pythonnet is not installed. Run:  pip install pythonnet"
        )

    # Change directory so DWSIM finds its thermo files
    os.chdir(dwsim_path)
    sys.path.append(dwsim_path)
    import clr as clr_inner

    required_dlls = [
        "DWSIM.Automation",
        "DWSIM.Interfaces",
        "DWSIM.GlobalSettings",
        "DWSIM.SharedClasses",
        "DWSIM.Thermodynamics",
        "DWSIM.UnitOperations",
        "DWSIM.Inspector",
        "System.Linq",
    ]
    for dll in required_dlls:
        try:
            clr_inner.AddReference(dll)
        except Exception:
            log.warning(f"Could not load DLL: {dll} – continuing anyway.")

    from DWSIM.Automation import Automation3  # type: ignore
    interf = Automation3()
    log.info("DWSIM Automation interface loaded successfully.")
    return interf


# ─────────────────────────────────────────────────────────────────────────────
# 2.  HELPER – UNIT CONVERSIONS, SAFE GETTERS & OBJECT CASTING
# ─────────────────────────────────────────────────────────────────────────────

def _safe(val, default=float("nan")):
    try:
        return float(val) if val is not None else default
    except Exception:
        return default

def _mol_frac(stream, compound_name: str) -> float:
    try:
        phases = stream.GetPhases() 
        comps = phases[0].Compounds
        if compound_name in comps:
            return _safe(comps[compound_name].MoleFraction.GetValueOrDefault())
        return float("nan")
    except Exception:
        try:
            phases = stream.Phases
            comps = phases[0].Compounds
            if compound_name in comps:
                return _safe(comps[compound_name].MoleFraction.GetValueOrDefault())
            return float("nan")
        except:
             return float("nan")

def _molar_flow(stream, compound_name: str) -> float:
    try:
        phases = stream.GetPhases() 
        comps = phases[0].Compounds
        if compound_name in comps:
            return _safe(comps[compound_name].MolarFlow.GetValueOrDefault())
        return float("nan")
    except Exception:
        try:
            phases = stream.Phases
            comps = phases[0].Compounds
            if compound_name in comps:
                return _safe(comps[compound_name].MolarFlow.GetValueOrDefault())
            return float("nan")
        except:
             return float("nan")

def _cast_to_stream(sim, stream_name: str):
    generic_obj = sim.GetFlowsheetSimulationObject(stream_name)
    generic_obj = generic_obj.GetAsObject() 
    return generic_obj

def _set_enum(obj, prop_name, int_val):
    import System
    try:
        current_val = getattr(obj, prop_name)
        enum_type = current_val.GetType()
        enum_val = System.Enum.ToObject(enum_type, int_val)
        setattr(obj, prop_name, enum_val)
    except Exception:
        setattr(obj, prop_name, int_val)

def _ot(name: str):
    from DWSIM.Interfaces.Enums.GraphicObjects import ObjectType
    import System
    try:
        return System.Enum.Parse(ObjectType, name, True)
    except Exception:
        names = System.Enum.GetNames(ObjectType)
        for n in names:
            if name.lower() in n.lower():
                return System.Enum.Parse(ObjectType, n, True)
        if name == "PFR":
            for n in names:
                if "plugflow" in n.lower() or "rct_pfr" in n.lower():
                    return System.Enum.Parse(ObjectType, n, True)
        raise ValueError(f"Could not find {name} in {list(names)}")

def _connect(sim, obj1_name: str, obj2_name: str, port1: int, port2: int):
    go1 = sim.GetFlowsheetSimulationObject(obj1_name).GraphicObject
    go2 = sim.GetFlowsheetSimulationObject(obj2_name).GraphicObject
    sim.ConnectObjects(go1, go2, port1, port2)


# ─────────────────────────────────────────────────────────────────────────────
# 3.  PART A – PFR SIMULATION
# ─────────────────────────────────────────────────────────────────────────────

def build_pfr_flowsheet(interf, volume_m3: float, feed_temp_K: float):
    import System 
    from System.Collections.Generic import Dictionary

    sim = interf.CreateFlowsheet()

    pp_name = "Peng-Robinson (PR)"
    sim.CreateAndAddPropertyPackage(pp_name)

    for c in ["n-Pentane", "Isopentane"]:
        sim.AddCompound(c)

    sim.AddObject(_ot("MaterialStream"), 100, 300, "FEED")
    sim.AddObject(_ot("MaterialStream"), 600, 300, "PRODUCT")
    
    try:
        pfr_type = _ot("RCT_PFR")
    except Exception:
        pfr_type = _ot("PFR")
        
    sim.AddObject(pfr_type, 350, 300, "PFR1")

    feed_obj = _cast_to_stream(sim, "FEED")
    composition = System.Array[float]([1.0, 0.0])
    feed_obj.SetOverallComposition(composition)           
    feed_obj.SetTemperature(feed_temp_K)
    feed_obj.SetPressure(101325.0)                        
    feed_obj.SetMassFlow(1.0)                             
    _set_enum(feed_obj, "SpecType", 0)                               

    pfr_obj = sim.GetFlowsheetSimulationObject("PFR1").GetAsObject()
    pfr_obj.Volume = volume_m3
    
    try:
        pfr_obj.Isothermal = True
        pfr_obj.IsothermalTemperature = feed_temp_K          
    except Exception:
        pass

    comps = Dictionary[System.String, System.Double]()
    comps.Add("n-Pentane", -1.0)
    comps.Add("Isopentane", 1.0)

    dorders = Dictionary[System.String, System.Double]()
    dorders.Add("n-Pentane", 1.0)
    dorders.Add("Isopentane", 0.0)

    rorders = Dictionary[System.String, System.Double]()
    rorders.Add("n-Pentane", 0.0)
    rorders.Add("Isopentane", 0.0)

    kr1 = sim.CreateKineticReaction(
        "nC5_isoC5", "Isomerization", 
        comps, dorders, rorders, 
        "n-Pentane", "Mixture", 
        "Molar Concentration", "mol/m3", "mol/[m3.s]", 
        1.0e6, 80000.0, 0.0, 0.0, "", ""
    )

    sim.AddReaction(kr1)
    sim.AddReactionToSet(kr1.ID, "DefaultSet", True, 0)
    pfr_obj.ReactionSet = "DefaultSet"

    _connect(sim, "FEED", "PFR1", -1, -1)
    _connect(sim, "PFR1", "PRODUCT", -1, -1)

    return sim


def extract_pfr_results(interf, sim, volume_m3: float, feed_temp_K: float) -> dict:
    result = {
        "part": "PFR",
        "volume_m3": volume_m3,
        "feed_temp_K": feed_temp_K,
        "success": False,
        "error": "",
        "conversion": float("nan"),
        "outlet_nC5_molflow": float("nan"),
        "outlet_iC5_molflow": float("nan"),
        "heat_duty_W": float("nan"),
        "outlet_temp_K": float("nan"),
    }
    try:
        if hasattr(interf, 'CalculateFlowsheet2'):
            interf.CalculateFlowsheet2(sim)
        else:
            sim.RequestCalculation()

        prod_obj = _cast_to_stream(sim, "PRODUCT")
        pfr_obj  = sim.GetFlowsheetSimulationObject("PFR1").GetAsObject()
        feed_obj = _cast_to_stream(sim, "FEED")

        f_nC5_in  = _molar_flow(feed_obj, "n-Pentane")
        f_nC5_out = _molar_flow(prod_obj, "n-Pentane")
        f_iC5_out = _molar_flow(prod_obj, "Isopentane")

        conversion = (f_nC5_in - f_nC5_out) / f_nC5_in if f_nC5_in > 0 else float("nan")

        result.update({
            "success": True,
            "conversion": round(conversion, 6),
            "outlet_nC5_molflow": round(f_nC5_out, 6),
            "outlet_iC5_molflow": round(f_iC5_out, 6),
            "heat_duty_W": _safe(getattr(pfr_obj, "DeltaQ", None)),
            "outlet_temp_K": _safe(prod_obj.GetTemperature()),
        })
    except Exception as e:
        result["error"] = str(e)
        log.warning(f"  PFR case V={volume_m3} T={feed_temp_K} FAILED: {e}")

    return result


# ─────────────────────────────────────────────────────────────────────────────
# 4.  PART B – DISTILLATION COLUMN SIMULATION
# ─────────────────────────────────────────────────────────────────────────────

def build_distil_flowsheet(interf, n_stages: int, feed_stage: int,
                            reflux_ratio: float, distillate_rate: float):
    import System 

    sim = interf.CreateFlowsheet()

    pp_name = "Peng-Robinson (PR)"
    sim.CreateAndAddPropertyPackage(pp_name)

    for c in ["n-Pentane", "Isopentane"]:
        sim.AddCompound(c)

    sim.AddObject(_ot("MaterialStream"), 100, 300, "FEED_D")
    sim.AddObject(_ot("MaterialStream"), 600, 200, "DISTILLATE")
    sim.AddObject(_ot("MaterialStream"), 600, 400, "BOTTOMS")
    sim.AddObject(_ot("DistillationColumn"), 350, 300, "COL1")

    feed_obj = _cast_to_stream(sim, "FEED_D")
    composition = System.Array[float]([0.5, 0.5])
    feed_obj.SetOverallComposition(composition)            
    feed_obj.SetTemperature(310.0)                        
    feed_obj.SetPressure(202650.0)                        
    feed_obj.SetMolarFlow(100.0)                          
    _set_enum(feed_obj, "SpecType", 0)

    col_obj = sim.GetFlowsheetSimulationObject("COL1").GetAsObject()
    col_obj.NumberOfStages = n_stages
    col_obj.FeedStage = feed_stage
    col_obj.RefluxRatio = reflux_ratio
    col_obj.DistillateFlowSpec = distillate_rate * 100.0  

    _set_enum(col_obj, "CondenserType", 0)    
    _set_enum(col_obj, "ReboilerType", 0)    

    _connect(sim, "FEED_D", "COL1", -1, 0)   
    _connect(sim, "COL1", "DISTILLATE", 0, -1)   
    _connect(sim, "COL1", "BOTTOMS", 1, -1)   

    return sim


def extract_distil_results(interf, sim, n_stages, feed_stage, reflux_ratio,
                            distillate_rate) -> dict:
    result = {
        "part": "Distillation",
        "n_stages": n_stages,
        "feed_stage": feed_stage,
        "reflux_ratio": reflux_ratio,
        "distillate_rate": distillate_rate,
        "success": False,
        "error": "",
        "distillate_iC5_purity": float("nan"),
        "bottoms_nC5_purity": float("nan"),
        "condenser_duty_W": float("nan"),
        "reboiler_duty_W": float("nan"),
    }
    try:
        if hasattr(interf, 'CalculateFlowsheet2'):
            interf.CalculateFlowsheet2(sim)
        else:
            sim.RequestCalculation()

        dist_obj = _cast_to_stream(sim, "DISTILLATE")
        bott_obj = _cast_to_stream(sim, "BOTTOMS")
        col_obj  = sim.GetFlowsheetSimulationObject("COL1").GetAsObject()

        result.update({
            "success": True,
            "distillate_iC5_purity": round(_mol_frac(dist_obj, "Isopentane"), 6),
            "bottoms_nC5_purity":    round(_mol_frac(bott_obj, "n-Pentane"), 6),
            "condenser_duty_W": _safe(getattr(col_obj, "CondenserDuty", None)),
            "reboiler_duty_W":  _safe(getattr(col_obj, "ReboilerDuty",  None)),
        })
    except Exception as e:
        result["error"] = str(e)
        log.warning(f"  Distil case N={n_stages} R={reflux_ratio} FAILED: {e}")

    return result


# ─────────────────────────────────────────────────────────────────────────────
# 5.  PARAMETRIC SWEEPS CONFIG
# ─────────────────────────────────────────────────────────────────────────────

PFR_VOLUMES   = [0.5, 1.0, 2.0, 5.0, 10.0]      
PFR_TEMPS     = [350.0, 380.0, 410.0, 440.0]     

DISTIL_STAGES  = [10, 15, 20, 25]
DISTIL_REFLUX  = [1.0, 1.5, 2.0, 3.0]
DISTIL_FEED_STAGE      = 8       
DISTIL_DISTILLATE_RATE = 0.50   


# ─────────────────────────────────────────────────────────────────────────────
# 6.  CSV OUTPUT
# ─────────────────────────────────────────────────────────────────────────────

CSV_FIELDNAMES = [
    "timestamp", "part", "case_id",
    "volume_m3", "feed_temp_K",
    "n_stages", "feed_stage", "reflux_ratio", "distillate_rate",
    "success", "error",
    "conversion", "outlet_nC5_molflow", "outlet_iC5_molflow",
    "heat_duty_W", "outlet_temp_K",
    "distillate_iC5_purity", "bottoms_nC5_purity",
    "condenser_duty_W", "reboiler_duty_W",
]

def write_csv(rows: list, filename: str = "results.csv"):
    # Guarantee the file goes to the original script directory
    path = os.path.join(SCRIPT_DIR, filename)
    with open(path, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_FIELDNAMES, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            full_row = {k: "" for k in CSV_FIELDNAMES}
            full_row.update(row)
            writer.writerow(full_row)
    log.info(f"Results written to {path}  ({len(rows)} rows)")


# ─────────────────────────────────────────────────────────────────────────────
# 7.  OPTIONAL PLOTS
# ─────────────────────────────────────────────────────────────────────────────

def make_plots(all_results: list):
    if not PLOT_AVAILABLE:
        log.info("matplotlib not available – skipping plots.")
        return

    pfr_rows   = [r for r in all_results if r["part"] == "PFR"          and r["success"]]
    dist_rows  = [r for r in all_results if r["part"] == "Distillation" and r["success"]]

    if pfr_rows:
        fig, axes = plt.subplots(1, 2, figsize=(12, 5))
        fig.suptitle("PFR Parametric Sweep – n-C5 Isomerization", fontsize=13)

        temps = sorted(set(r["feed_temp_K"] for r in pfr_rows))
        colors = plt.cm.plasma([i / max(len(temps) - 1, 1) for i in range(len(temps))])

        for temp, color in zip(temps, colors):
            subset = sorted([r for r in pfr_rows if r["feed_temp_K"] == temp], key=lambda x: x["volume_m3"])
            vols = [r["volume_m3"]  for r in subset]
            conv = [r["conversion"] for r in subset]
            axes[0].plot(vols, conv, marker="o", label=f"T={int(temp)} K", color=color)

        axes[0].set_xlabel("Reactor Volume (m³)"); axes[0].set_ylabel("n-C5 Conversion"); axes[0].legend(); axes[0].grid(True, alpha=0.3)

        vols_u = sorted(set(r["volume_m3"] for r in pfr_rows))
        colors2 = plt.cm.viridis([i / max(len(vols_u) - 1, 1) for i in range(len(vols_u))])
        for vol, color in zip(vols_u, colors2):
            subset = sorted([r for r in pfr_rows if r["volume_m3"] == vol], key=lambda x: x["feed_temp_K"])
            temps2 = [r["feed_temp_K"] for r in subset]
            conv   = [r["conversion"]  for r in subset]
            axes[1].plot(temps2, conv, marker="s", label=f"V={vol} m³", color=color)

        axes[1].set_xlabel("Feed Temperature (K)"); axes[1].set_ylabel("n-C5 Conversion"); axes[1].legend(); axes[1].grid(True, alpha=0.3)
        plt.tight_layout()
        
        # Guarantee output path
        pfr_path = os.path.join(SCRIPT_DIR, "pfr_sweep.png")
        plt.savefig(pfr_path, dpi=150, bbox_inches="tight")
        plt.close()
        log.info(f"Saved {pfr_path}")

    if dist_rows:
        fig, axes = plt.subplots(1, 2, figsize=(12, 5))
        fig.suptitle("Distillation Parametric Sweep – n-C5 / i-C5 Separation", fontsize=13)

        stages_u = sorted(set(r["n_stages"] for r in dist_rows))
        colors = plt.cm.coolwarm([i / max(len(stages_u) - 1, 1) for i in range(len(stages_u))])

        for ns, color in zip(stages_u, colors):
            subset = sorted([r for r in dist_rows if r["n_stages"] == ns], key=lambda x: x["reflux_ratio"])
            rr   = [r["reflux_ratio"]           for r in subset]
            pur  = [r["distillate_iC5_purity"]  for r in subset]
            axes[0].plot(rr, pur, marker="o", label=f"N={ns} stages", color=color)

        axes[0].set_xlabel("Reflux Ratio (L/D)"); axes[0].set_ylabel("i-C5 Purity in Distillate (mol frac)"); axes[0].legend(); axes[0].grid(True, alpha=0.3)

        for ns, color in zip(stages_u, colors):
            subset = sorted([r for r in dist_rows if r["n_stages"] == ns], key=lambda x: x["reflux_ratio"])
            rr   = [r["reflux_ratio"]        for r in subset]
            pur  = [r["bottoms_nC5_purity"]  for r in subset]
            axes[1].plot(rr, pur, marker="s", label=f"N={ns} stages", color=color)

        axes[1].set_xlabel("Reflux Ratio (L/D)"); axes[1].set_ylabel("n-C5 Purity in Bottoms (mol frac)"); axes[1].legend(); axes[1].grid(True, alpha=0.3)
        plt.tight_layout()
        
        # Guarantee output path
        distil_path = os.path.join(SCRIPT_DIR, "distil_sweep.png")
        plt.savefig(distil_path, dpi=150, bbox_inches="tight")
        plt.close()
        log.info(f"Saved {distil_path}")


# ─────────────────────────────────────────────────────────────────────────────
# 8.  MAIN EXECUTION
# ─────────────────────────────────────────────────────────────────────────────

def main():
    log.info("=" * 65)
    log.info("FOSSEE DWSIM Automation Screening Task – Starting")
    log.info(f"Run timestamp: {datetime.now().isoformat()}")
    log.info("=" * 65)

    dwsim_path = find_dwsim_path()
    interf = load_dwsim(dwsim_path)

    all_results = []
    case_id = 0

    log.info("\n--- Part A/C: PFR Parametric Sweep ---")
    pfr_combos = list(itertools.product(PFR_VOLUMES, PFR_TEMPS))
    log.info(f"  Total PFR cases: {len(pfr_combos)}")

    for volume, temp in pfr_combos:
        case_id += 1
        log.info(f"  Case {case_id}: V={volume} m³, T={temp} K")
        try:
            sim = build_pfr_flowsheet(interf, volume, temp)
            row = extract_pfr_results(interf, sim, volume, temp)
        except Exception as exc:
            log.error(f"  Case {case_id} build/extract ERROR: {exc}\n{traceback.format_exc()}")
            row = {
                "part": "PFR", "volume_m3": volume, "feed_temp_K": temp,
                "success": False, "error": str(exc),
            }
        row["case_id"]   = case_id
        row["timestamp"] = datetime.now().isoformat()
        all_results.append(row)

    log.info("\n--- Part B/C: Distillation Parametric Sweep ---")
    dist_combos = list(itertools.product(DISTIL_STAGES, DISTIL_REFLUX))
    log.info(f"  Total Distillation cases: {len(dist_combos)}")

    for n_stages, reflux in dist_combos:
        feed_stage = min(DISTIL_FEED_STAGE, n_stages - 2)
        case_id += 1
        log.info(f"  Case {case_id}: N={n_stages}, feed_stage={feed_stage}, R={reflux}")
        try:
            sim = build_distil_flowsheet(
                interf, n_stages, feed_stage, reflux, DISTIL_DISTILLATE_RATE
            )
            row = extract_distil_results(
                interf, sim, n_stages, feed_stage, reflux, DISTIL_DISTILLATE_RATE
            )
        except Exception as exc:
            log.error(f"  Case {case_id} build/extract ERROR: {exc}\n{traceback.format_exc()}")
            row = {
                "part": "Distillation", "n_stages": n_stages,
                "feed_stage": feed_stage, "reflux_ratio": reflux,
                "distillate_rate": DISTIL_DISTILLATE_RATE,
                "success": False, "error": str(exc),
            }
        row["case_id"]   = case_id
        row["timestamp"] = datetime.now().isoformat()
        all_results.append(row)

    write_csv(all_results)
    make_plots(all_results)

    total   = len(all_results)
    passed  = sum(1 for r in all_results if r.get("success"))
    failed  = total - passed
    log.info("\n" + "=" * 65)
    log.info(f"SUMMARY:  {total} cases run |  {passed} succeeded |  {failed} failed")
    log.info("Outputs:  results.csv, run_screening.log")
    if PLOT_AVAILABLE:
        log.info("          pfr_sweep.png, distil_sweep.png")
    log.info("=" * 65)

if __name__ == "__main__":
    main()