"""
FOSSEE Summer Fellowship 2026 - DWSIM Python Automation Screening Task
Task 1: Python Automation of DWSIM
"""

import sys
import os
import csv
import traceback
import itertools
import logging
from datetime import datetime

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
        logging.FileHandler("run_screening.log", mode="w"),
    ],
)
log = logging.getLogger(__name__)

# ─────────────────────────────────────────────────────────────────────────────
# 1.  DWSIM AUTOMATION BOOTSTRAP (FIXED BRIDGE)
# ─────────────────────────────────────────────────────────────────────────────

def find_dwsim_path():
    """Return the path to the DWSIM installation directory."""
    candidates = [
        r"C:\Users\athar\AppData\Local\DWSIM",  # Your confirmed path
        r"C:\Users\Public\DWSIM8",
        r"C:\Program Files\DWSIM8",
        r"C:\Program Files\DWSIM"
    ]
    for p in candidates:
        if p and os.path.isdir(p):
            log.info(f"DWSIM found at: {p}")
            return p
    raise EnvironmentError("DWSIM installation not found.")


def load_dwsim(dwsim_path: str):
    """
    Bootstrap pythonnet and load DWSIM assemblies using absolute paths.
    """
    try:
        import clr  
    except ImportError:
        raise ImportError("pythonnet is not installed. Run: pip install pythonnet")

    sys.path.append(dwsim_path)
    
    required_dlls = [
        "DWSIM.Automation.dll",
        "DWSIM.Interfaces.dll",
        "DWSIM.GlobalSettings.dll",
        "DWSIM.SharedClasses.dll",
        "DWSIM.Thermodynamics.dll",
        "DWSIM.UnitOperations.dll",
        "DWSIM.Inspector.dll"
    ]
    
    for dll in required_dlls:
        dll_file_path = os.path.join(dwsim_path, dll)
        if os.path.exists(dll_file_path):
            try:
                clr.AddReference(dll_file_path)
            except Exception as e:
                log.warning(f"Could not load DLL {dll}: {e}")
        else:
            log.warning(f"WARNING: File not found - {dll_file_path}")

    try:
        clr.AddReference("System.Linq")
    except:
        pass

    try:
        from DWSIM.Automation import Automation3
        interf = Automation3()
        log.info("Loaded Automation3 interface.")
    except ImportError:
        try:
            from DWSIM.Automation import Automation2
            interf = Automation2()
            log.info("Loaded Automation2 interface.")
        except ImportError:
            from DWSIM.Automation import Automation
            interf = Automation()
            log.info("Loaded base Automation interface.")
            
    log.info("DWSIM Automation interface loaded successfully.")
    return interf


# ─────────────────────────────────────────────────────────────────────────────
# 2.  HELPER – UNIT CONVERSIONS & SAFE GETTERS
# ─────────────────────────────────────────────────────────────────────────────

def _safe(val, default=float("nan")):
    try:
        return float(val) if val is not None else default
    except Exception:
        return default

def _mol_frac(stream, compound_name: str) -> float:
    try:
        phases = stream.Phases
        comps = phases[0].Compounds
        if compound_name in comps:
            return _safe(comps[compound_name].MoleFraction.GetValueOrDefault())
        return float("nan")
    except Exception:
        return float("nan")

def _molar_flow(stream, compound_name: str) -> float:
    try:
        phases = stream.Phases
        comps = phases[0].Compounds
        if compound_name in comps:
            return _safe(comps[compound_name].MolarFlow.GetValueOrDefault())
        return float("nan")
    except Exception:
        return float("nan")


# ─────────────────────────────────────────────────────────────────────────────
# 3.  PART A – PFR SIMULATION
# ─────────────────────────────────────────────────────────────────────────────

def build_pfr_flowsheet(interf, volume_m3: float, feed_temp_K: float):
    sim = interf.CreateFlowsheet()
    sim.AddPropertyPackage("Peng-Robinson")
    
    for c in ["n-Pentane", "Isopentane"]:
        sim.AddCompound(c)

    feed = sim.AddObject("MaterialStream", 100, 300, "FEED")
    prod = sim.AddObject("MaterialStream", 600, 300, "PRODUCT")

    feed_obj = sim.GetFlowsheetSimulationObject("FEED")
    feed_obj.SetOverallComposition([1.0, 0.0])           
    feed_obj.SetTemperature(feed_temp_K)
    feed_obj.SetPressure(101325.0)                        
    feed_obj.SetMassFlow(1.0)                             
    feed_obj.SpecType = 0                                 

    pfr = sim.AddObject("PFR", 350, 300, "PFR1")
    pfr_obj = sim.GetFlowsheetSimulationObject("PFR1")

    pfr_obj.Volume = volume_m3
    pfr_obj.Isothermal = True
    pfr_obj.IsothermalTemperature = feed_temp_K          

    rxn = pfr_obj.AddReaction()
    rxn.Name = "nC5_isoC5"
    rxn.ReactionType = 0                                  
    rxn.KineticExpression = 0                             

    rxn.AddComponent("n-Pentane", -1.0)
    rxn.AddComponent("Isopentane", 1.0)

    rxn.PreExponentialFactor = 1.0e6                      
    rxn.ActivationEnergy = 80000.0                        
    rxn.ReactionOrder = 1.0

    sim.ConnectObjects(feed, pfr, -1, -1)
    sim.ConnectObjects(pfr, prod, -1, -1)

    return sim

def extract_pfr_results(sim, volume_m3: float, feed_temp_K: float) -> dict:
    result = {
        "part": "PFR", "volume_m3": volume_m3, "feed_temp_K": feed_temp_K,
        "success": False, "error": "", "conversion": float("nan"),
        "outlet_nC5_molflow": float("nan"), "outlet_iC5_molflow": float("nan"),
        "heat_duty_W": float("nan"), "outlet_temp_K": float("nan"),
    }
    try:
        sim.SolveFlowsheet()
        prod_obj = sim.GetFlowsheetSimulationObject("PRODUCT")
        pfr_obj  = sim.GetFlowsheetSimulationObject("PFR1")
        feed_obj = sim.GetFlowsheetSimulationObject("FEED")

        f_nC5_in  = _molar_flow(feed_obj, "n-Pentane")
        f_nC5_out = _molar_flow(prod_obj, "n-Pentane")
        f_iC5_out = _molar_flow(prod_obj, "Isopentane")

        conversion = (f_nC5_in - f_nC5_out) / f_nC5_in if f_nC5_in > 0 else float("nan")

        result.update({
            "success": True, "conversion": round(conversion, 6),
            "outlet_nC5_molflow": round(f_nC5_out, 6), "outlet_iC5_molflow": round(f_iC5_out, 6),
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

def build_distil_flowsheet(interf, n_stages: int, feed_stage: int, reflux_ratio: float, distillate_rate: float):
    sim = interf.CreateFlowsheet()
    sim.AddPropertyPackage("Peng-Robinson")
    for c in ["n-Pentane", "Isopentane"]:
        sim.AddCompound(c)

    feed = sim.AddObject("MaterialStream", 100, 300, "FEED_D")
    dist = sim.AddObject("MaterialStream", 600, 200, "DISTILLATE")
    bott = sim.AddObject("MaterialStream", 600, 400, "BOTTOMS")

    feed_obj = sim.GetFlowsheetSimulationObject("FEED_D")
    feed_obj.SetOverallComposition([0.5, 0.5])            
    feed_obj.SetTemperature(310.0)                        
    feed_obj.SetPressure(202650.0)                        
    feed_obj.SetMolarFlow(100.0)                          
    feed_obj.SpecType = 0

    col = sim.AddObject("DistillationColumn", 350, 300, "COL1")
    col_obj = sim.GetFlowsheetSimulationObject("COL1")

    col_obj.NumberOfStages = n_stages
    col_obj.FeedStage = feed_stage
    col_obj.RefluxRatio = reflux_ratio
    col_obj.DistillateFlowSpec = distillate_rate * 100.0  

    col_obj.CondenserType = 0    
    col_obj.ReboilerType  = 0    

    sim.ConnectObjects(feed, col, -1, 0)   
    sim.ConnectObjects(col, dist, 0, -1)   
    sim.ConnectObjects(col, bott, 1, -1)   

    return sim

def extract_distil_results(sim, n_stages, feed_stage, reflux_ratio, distillate_rate) -> dict:
    result = {
        "part": "Distillation", "n_stages": n_stages, "feed_stage": feed_stage,
        "reflux_ratio": reflux_ratio, "distillate_rate": distillate_rate,
        "success": False, "error": "", "distillate_iC5_purity": float("nan"),
        "bottoms_nC5_purity": float("nan"), "condenser_duty_W": float("nan"),
        "reboiler_duty_W": float("nan"),
    }
    try:
        sim.SolveFlowsheet()
        dist_obj = sim.GetFlowsheetSimulationObject("DISTILLATE")
        bott_obj = sim.GetFlowsheetSimulationObject("BOTTOMS")
        col_obj  = sim.GetFlowsheetSimulationObject("COL1")

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
# 5.  PART C – PARAMETRIC SWEEPS
# ─────────────────────────────────────────────────────────────────────────────

PFR_VOLUMES   = [0.5, 1.0, 2.0, 5.0, 10.0]      
PFR_TEMPS     = [350.0, 380.0, 410.0, 440.0]     

DISTIL_STAGES  = [10, 15, 20, 25]
DISTIL_REFLUX  = [1.0, 1.5, 2.0, 3.0]
DISTIL_FEED_STAGE      = 8       
DISTIL_DISTILLATE_RATE = 0.50   

# ─────────────────────────────────────────────────────────────────────────────
# 6.  CSV OUTPUT & OPTIONAL PLOTS
# ─────────────────────────────────────────────────────────────────────────────

CSV_FIELDNAMES = [
    "timestamp", "part", "case_id", "volume_m3", "feed_temp_K",
    "n_stages", "feed_stage", "reflux_ratio", "distillate_rate",
    "success", "error", "conversion", "outlet_nC5_molflow", "outlet_iC5_molflow",
    "heat_duty_W", "outlet_temp_K", "distillate_iC5_purity", "bottoms_nC5_purity",
    "condenser_duty_W", "reboiler_duty_W",
]

def write_csv(rows: list, path: str = "results.csv"):
    with open(path, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_FIELDNAMES, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            full_row = {k: "" for k in CSV_FIELDNAMES}
            full_row.update(row)
            writer.writerow(full_row)
    log.info(f"Results written to {path}  ({len(rows)} rows)")

def make_plots(all_results: list):
    if not PLOT_AVAILABLE: return
    pfr_rows   = [r for r in all_results if r["part"] == "PFR" and r["success"]]
    if pfr_rows:
        fig, axes = plt.subplots(1, 2, figsize=(12, 5))
        temps = sorted(set(r["feed_temp_K"] for r in pfr_rows))
        for temp in temps:
            subset = sorted([r for r in pfr_rows if r["feed_temp_K"] == temp], key=lambda x: x["volume_m3"])
            axes[0].plot([r["volume_m3"] for r in subset], [r["conversion"] for r in subset], marker="o", label=f"T={int(temp)} K")
        axes[0].set_xlabel("Reactor Volume (m³)"); axes[0].set_ylabel("Conversion"); axes[0].legend()
        plt.tight_layout()
        plt.savefig("pfr_sweep.png", dpi=150)
        plt.close()
        log.info("Saved pfr_sweep.png")

# ─────────────────────────────────────────────────────────────────────────────
# 8.  MAIN EXECUTION
# ─────────────────────────────────────────────────────────────────────────────

def main():
    log.info("=" * 65)
    log.info("FOSSEE DWSIM Automation Screening Task – Starting")
    log.info("=" * 65)
    
    dwsim_path = find_dwsim_path()
    interf = load_dwsim(dwsim_path)
    all_results = []
    case_id = 0

    log.info("\n--- Part A/C: PFR Parametric Sweep ---")
    for volume, temp in itertools.product(PFR_VOLUMES, PFR_TEMPS):
        case_id += 1
        log.info(f"  Case {case_id}: V={volume} m³, T={temp} K")
        try:
            sim = build_pfr_flowsheet(interf, volume, temp)
            row = extract_pfr_results(sim, volume, temp)
        except Exception as exc:
            row = {"part": "PFR", "volume_m3": volume, "feed_temp_K": temp, "success": False, "error": str(exc)}
        row["case_id"] = case_id
        row["timestamp"] = datetime.now().isoformat()
        all_results.append(row)

    log.info("\n--- Part B/C: Distillation Parametric Sweep ---")
    for n_stages, reflux in itertools.product(DISTIL_STAGES, DISTIL_REFLUX):
        feed_stage = min(DISTIL_FEED_STAGE, n_stages - 2)
        case_id += 1
        log.info(f"  Case {case_id}: N={n_stages}, feed_stage={feed_stage}, R={reflux}")
        try:
            sim = build_distil_flowsheet(interf, n_stages, feed_stage, reflux, DISTIL_DISTILLATE_RATE)
            row = extract_distil_results(sim, n_stages, feed_stage, reflux, DISTIL_DISTILLATE_RATE)
        except Exception as exc:
            row = {"part": "Distillation", "n_stages": n_stages, "reflux_ratio": reflux, "success": False, "error": str(exc)}
        row["case_id"] = case_id
        row["timestamp"] = datetime.now().isoformat()
        all_results.append(row)

    write_csv(all_results)
    make_plots(all_results)
    
    passed = sum(1 for r in all_results if r.get("success"))
    log.info(f"\nSUMMARY: {len(all_results)} cases | {passed} succeeded")

if __name__ == "__main__":
    main()