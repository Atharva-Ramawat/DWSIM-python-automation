"""
FOSSEE Summer Fellowship 2026 - DWSIM Python Automation Screening Task
Task 1: Python Automation of DWSIM

Author: [Your Name]
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
        os.environ.get("DWSIM_HOME", ""),
        # Linux / macOS via Mono
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
        import clr  # noqa: F401  (pythonnet)
    except ImportError:
        raise ImportError(
            "pythonnet is not installed. Run:  pip install pythonnet"
        )

    # Add DWSIM assemblies to the CLR search path
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
# 2.  HELPER – UNIT CONVERSIONS & SAFE GETTERS
# ─────────────────────────────────────────────────────────────────────────────

def _safe(val, default=float("nan")):
    """Return val, or default if val is None / raises."""
    try:
        return float(val) if val is not None else default
    except Exception:
        return default


def _mol_frac(stream, compound_name: str) -> float:
    """Return mole fraction of compound_name in stream (0–1)."""
    try:
        phases = stream.Phases
        # Phase 0 = overall mixture
        comps = phases[0].Compounds
        if compound_name in comps:
            return _safe(comps[compound_name].MoleFraction.GetValueOrDefault())
        return float("nan")
    except Exception:
        return float("nan")


def _molar_flow(stream, compound_name: str) -> float:
    """Return molar flow of compound_name [mol/s] in stream."""
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
    """
    Create a DWSIM flowsheet containing:
        Feed stream → PFR (isomerization n-C5 → i-C5) → Product stream

    Returns the flowsheet object (IFlowsheet).
    """
    from DWSIM.Automation import Automation3          # type: ignore
    from System import String                          # type: ignore  # noqa

    sim = interf.CreateFlowsheet()

    # ── Property Package: Peng-Robinson ──────────────────────────────────────
    pp_name = "Peng-Robinson"
    sim.AddPropertyPackage(pp_name)

    # ── Compounds ─────────────────────────────────────────────────────────────
    for c in ["n-Pentane", "Isopentane"]:
        sim.AddCompound(c)

    # ── Material Streams ──────────────────────────────────────────────────────
    feed = sim.AddObject("MaterialStream", 100, 300, "FEED")
    prod = sim.AddObject("MaterialStream", 600, 300, "PRODUCT")

    # Feed conditions
    feed_obj = sim.GetFlowsheetSimulationObject("FEED")
    feed_obj.SetOverallComposition([1.0, 0.0])           # pure n-C5
    feed_obj.SetTemperature(feed_temp_K)
    feed_obj.SetPressure(101325.0)                        # 1 atm [Pa]
    feed_obj.SetMassFlow(1.0)                             # 1 kg/s basis
    feed_obj.SpecType = 0                                 # T, P spec

    # ── PFR Unit Operation ────────────────────────────────────────────────────
    pfr = sim.AddObject("PFR", 350, 300, "PFR1")
    pfr_obj = sim.GetFlowsheetSimulationObject("PFR1")

    pfr_obj.Volume = volume_m3
    pfr_obj.Isothermal = True
    pfr_obj.IsothermalTemperature = feed_temp_K          # same as feed

    # Kinetics: first-order isomerization  r = k·C_nC5
    # k = A·exp(-Ea/RT),  A=1e6 s⁻¹,  Ea=80 kJ/mol  (literature-based estimate)
    rxn = pfr_obj.AddReaction()
    rxn.Name = "nC5_isoC5"
    rxn.ReactionType = 0                                  # kinetic
    rxn.KineticExpression = 0                             # power law

    # Reactants / products
    rxn.AddComponent("n-Pentane", -1.0)
    rxn.AddComponent("Isopentane", 1.0)

    rxn.PreExponentialFactor = 1.0e6                      # A  [1/s]
    rxn.ActivationEnergy = 80000.0                        # Ea [J/mol]
    rxn.ReactionOrder = 1.0

    # ── Connections ──────────────────────────────────────────────────────────
    sim.ConnectObjects(feed, pfr, -1, -1)
    sim.ConnectObjects(pfr, prod, -1, -1)

    return sim


def extract_pfr_results(sim, volume_m3: float, feed_temp_K: float) -> dict:
    """Run the PFR flowsheet and extract KPIs."""
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
        sim.SolveFlowsheet()

        prod_obj = sim.GetFlowsheetSimulationObject("PRODUCT")
        pfr_obj  = sim.GetFlowsheetSimulationObject("PFR1")
        feed_obj = sim.GetFlowsheetSimulationObject("FEED")

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
    """
    Create a DWSIM flowsheet containing:
        Feed → Distillation Column → Distillate + Bottoms

    n_stages        – total equilibrium stages (including condenser & reboiler)
    feed_stage      – stage number for feed entry (1-indexed from top)
    reflux_ratio    – L/D (external reflux ratio)
    distillate_rate – molar distillate-to-feed ratio [mol/mol]
    """
    sim = interf.CreateFlowsheet()

    pp_name = "Peng-Robinson"
    sim.AddPropertyPackage(pp_name)

    for c in ["n-Pentane", "Isopentane"]:
        sim.AddCompound(c)

    # Streams
    feed = sim.AddObject("MaterialStream", 100, 300, "FEED_D")
    dist = sim.AddObject("MaterialStream", 600, 200, "DISTILLATE")
    bott = sim.AddObject("MaterialStream", 600, 400, "BOTTOMS")

    feed_obj = sim.GetFlowsheetSimulationObject("FEED_D")
    feed_obj.SetOverallComposition([0.5, 0.5])            # equimolar n-C5/i-C5
    feed_obj.SetTemperature(310.0)                        # K (slightly subcooled)
    feed_obj.SetPressure(202650.0)                        # 2 atm [Pa]
    feed_obj.SetMolarFlow(100.0)                          # 100 mol/s basis
    feed_obj.SpecType = 0

    # Distillation column
    col = sim.AddObject("DistillationColumn", 350, 300, "COL1")
    col_obj = sim.GetFlowsheetSimulationObject("COL1")

    col_obj.NumberOfStages = n_stages
    col_obj.FeedStage = feed_stage
    col_obj.RefluxRatio = reflux_ratio
    # 4th spec: distillate-to-feed molar ratio
    col_obj.DistillateFlowSpec = distillate_rate * 100.0  # mol/s

    col_obj.CondenserType = 0    # total condenser
    col_obj.ReboilerType  = 0    # kettle reboiler

    # Connections
    sim.ConnectObjects(feed, col, -1, 0)   # feed → column feed inlet
    sim.ConnectObjects(col, dist, 0, -1)   # column distillate → stream
    sim.ConnectObjects(col, bott, 1, -1)   # column bottoms → stream

    return sim


def extract_distil_results(sim, n_stages, feed_stage, reflux_ratio,
                            distillate_rate) -> dict:
    """Run the distillation flowsheet and extract KPIs."""
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

# PFR sweep ranges
PFR_VOLUMES   = [0.5, 1.0, 2.0, 5.0, 10.0]      # m³
PFR_TEMPS     = [350.0, 380.0, 410.0, 440.0]     # K

# Distillation sweep ranges
DISTIL_STAGES  = [10, 15, 20, 25]
DISTIL_REFLUX  = [1.0, 1.5, 2.0, 3.0]

# Fixed distillation parameters (held constant during sweep)
DISTIL_FEED_STAGE      = 8       # feed stage (middle-ish)
DISTIL_DISTILLATE_RATE = 0.50   # 50 % of feed as distillate


# ─────────────────────────────────────────────────────────────────────────────
# 6.  CSV OUTPUT
# ─────────────────────────────────────────────────────────────────────────────

CSV_FIELDNAMES = [
    # Metadata
    "timestamp", "part", "case_id",
    # PFR sweep variables
    "volume_m3", "feed_temp_K",
    # Distillation sweep variables
    "n_stages", "feed_stage", "reflux_ratio", "distillate_rate",
    # Status
    "success", "error",
    # PFR KPIs
    "conversion", "outlet_nC5_molflow", "outlet_iC5_molflow",
    "heat_duty_W", "outlet_temp_K",
    # Distillation KPIs
    "distillate_iC5_purity", "bottoms_nC5_purity",
    "condenser_duty_W", "reboiler_duty_W",
]


def write_csv(rows: list, path: str = "results.csv"):
    with open(path, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_FIELDNAMES, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            # Fill missing keys with empty string
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

    # ── Plot 1: PFR Conversion vs Volume for each Temperature ────────────────
    if pfr_rows:
        fig, axes = plt.subplots(1, 2, figsize=(12, 5))
        fig.suptitle("PFR Parametric Sweep – n-C5 Isomerization", fontsize=13)

        temps = sorted(set(r["feed_temp_K"] for r in pfr_rows))
        colors = plt.cm.plasma([i / max(len(temps) - 1, 1) for i in range(len(temps))])

        for temp, color in zip(temps, colors):
            subset = sorted(
                [r for r in pfr_rows if r["feed_temp_K"] == temp],
                key=lambda x: x["volume_m3"],
            )
            vols = [r["volume_m3"]  for r in subset]
            conv = [r["conversion"] for r in subset]
            axes[0].plot(vols, conv, marker="o", label=f"T={int(temp)} K", color=color)

        axes[0].set_xlabel("Reactor Volume (m³)")
        axes[0].set_ylabel("n-C5 Conversion")
        axes[0].set_title("Conversion vs Volume")
        axes[0].legend()
        axes[0].grid(True, alpha=0.3)

        # Plot 2: Conversion vs Temperature for each Volume
        vols_u = sorted(set(r["volume_m3"] for r in pfr_rows))
        colors2 = plt.cm.viridis([i / max(len(vols_u) - 1, 1) for i in range(len(vols_u))])
        for vol, color in zip(vols_u, colors2):
            subset = sorted(
                [r for r in pfr_rows if r["volume_m3"] == vol],
                key=lambda x: x["feed_temp_K"],
            )
            temps2 = [r["feed_temp_K"] for r in subset]
            conv   = [r["conversion"]  for r in subset]
            axes[1].plot(temps2, conv, marker="s", label=f"V={vol} m³", color=color)

        axes[1].set_xlabel("Feed Temperature (K)")
        axes[1].set_ylabel("n-C5 Conversion")
        axes[1].set_title("Conversion vs Temperature")
        axes[1].legend()
        axes[1].grid(True, alpha=0.3)

        plt.tight_layout()
        plt.savefig("pfr_sweep.png", dpi=150, bbox_inches="tight")
        plt.close()
        log.info("Saved pfr_sweep.png")

    # ── Plot 3 & 4: Distillation purities ────────────────────────────────────
    if dist_rows:
        fig, axes = plt.subplots(1, 2, figsize=(12, 5))
        fig.suptitle("Distillation Parametric Sweep – n-C5 / i-C5 Separation", fontsize=13)

        stages_u = sorted(set(r["n_stages"] for r in dist_rows))
        colors = plt.cm.coolwarm([i / max(len(stages_u) - 1, 1) for i in range(len(stages_u))])

        for ns, color in zip(stages_u, colors):
            subset = sorted(
                [r for r in dist_rows if r["n_stages"] == ns],
                key=lambda x: x["reflux_ratio"],
            )
            rr   = [r["reflux_ratio"]           for r in subset]
            pur  = [r["distillate_iC5_purity"]  for r in subset]
            axes[0].plot(rr, pur, marker="o", label=f"N={ns} stages", color=color)

        axes[0].set_xlabel("Reflux Ratio (L/D)")
        axes[0].set_ylabel("i-C5 Purity in Distillate (mol frac)")
        axes[0].set_title("Distillate Purity vs Reflux Ratio")
        axes[0].legend()
        axes[0].grid(True, alpha=0.3)

        for ns, color in zip(stages_u, colors):
            subset = sorted(
                [r for r in dist_rows if r["n_stages"] == ns],
                key=lambda x: x["reflux_ratio"],
            )
            rr   = [r["reflux_ratio"]        for r in subset]
            pur  = [r["bottoms_nC5_purity"]  for r in subset]
            axes[1].plot(rr, pur, marker="s", label=f"N={ns} stages", color=color)

        axes[1].set_xlabel("Reflux Ratio (L/D)")
        axes[1].set_ylabel("n-C5 Purity in Bottoms (mol frac)")
        axes[1].set_title("Bottoms Purity vs Reflux Ratio")
        axes[1].legend()
        axes[1].grid(True, alpha=0.3)

        plt.tight_layout()
        plt.savefig("distil_sweep.png", dpi=150, bbox_inches="tight")
        plt.close()
        log.info("Saved distil_sweep.png")


# ─────────────────────────────────────────────────────────────────────────────
# 8.  MAIN EXECUTION
# ─────────────────────────────────────────────────────────────────────────────

def main():
    log.info("=" * 65)
    log.info("FOSSEE DWSIM Automation Screening Task – Starting")
    log.info(f"Run timestamp: {datetime.now().isoformat()}")
    log.info("=" * 65)

    # ── Bootstrap DWSIM ───────────────────────────────────────────────────────
    dwsim_path = find_dwsim_path()
    interf = load_dwsim(dwsim_path)

    all_results = []
    case_id = 0

    # ════════════════════════════════════════════════════════════════════════
    # PART A + C (PFR) – Parametric Sweep
    # ════════════════════════════════════════════════════════════════════════
    log.info("\n--- Part A/C: PFR Parametric Sweep ---")
    pfr_combos = list(itertools.product(PFR_VOLUMES, PFR_TEMPS))
    log.info(f"  Total PFR cases: {len(pfr_combos)}")

    for volume, temp in pfr_combos:
        case_id += 1
        log.info(f"  Case {case_id}: V={volume} m³, T={temp} K")
        try:
            sim = build_pfr_flowsheet(interf, volume, temp)
            row = extract_pfr_results(sim, volume, temp)
        except Exception as exc:
            log.error(f"  Case {case_id} build/extract ERROR: {exc}\n{traceback.format_exc()}")
            row = {
                "part": "PFR", "volume_m3": volume, "feed_temp_K": temp,
                "success": False, "error": str(exc),
            }
        row["case_id"]   = case_id
        row["timestamp"] = datetime.now().isoformat()
        all_results.append(row)

    # ════════════════════════════════════════════════════════════════════════
    # PART B + C (Distillation) – Parametric Sweep
    # ════════════════════════════════════════════════════════════════════════
    log.info("\n--- Part B/C: Distillation Parametric Sweep ---")
    dist_combos = list(itertools.product(DISTIL_STAGES, DISTIL_REFLUX))
    log.info(f"  Total Distillation cases: {len(dist_combos)}")

    for n_stages, reflux in dist_combos:
        # Ensure feed stage is within column (middle of rectifying section)
        feed_stage = min(DISTIL_FEED_STAGE, n_stages - 2)
        case_id += 1
        log.info(f"  Case {case_id}: N={n_stages}, feed_stage={feed_stage}, R={reflux}")
        try:
            sim = build_distil_flowsheet(
                interf, n_stages, feed_stage, reflux, DISTIL_DISTILLATE_RATE
            )
            row = extract_distil_results(
                sim, n_stages, feed_stage, reflux, DISTIL_DISTILLATE_RATE
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

    # ── Write CSV ─────────────────────────────────────────────────────────────
    write_csv(all_results)

    # ── Plots ─────────────────────────────────────────────────────────────────
    make_plots(all_results)

    # ── Summary ───────────────────────────────────────────────────────────────
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