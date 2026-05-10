# DWSIM Python Automation - FOSSEE Screening Task 2026

## 🚀 Project Overview
This project  provides a robust Python automation framework for **DWSIM**, built using the Automation3 interface and `pythonnet`. It was developed as a screening task for the FOSSEE Summer Fellowship 2026.

The suite programmatically constructs chemical flowsheets, performs multi-variable parametric sweeps, and extracts Key Performance Indicators (KPIs) into a structured CSV format—all without opening the DWSIM GUI.

## ✨ Key Features
* **Headless Execution:** Full background simulation using the DWSIM Automation API.
* **Parametric Sweeps:** * **Reactor (PFR):** Conversion analysis across multiple volumes and temperatures.
  * **Distillation:** Purity and energy duty analysis across various stages and reflux ratios.
* **Smart Enum Parsing:** Version-agnostic object creation using C# Reflection.
* **Data Export:** Automated generation of `results.csv` and visualization plots.

## 🛠️ Requirements
* DWSIM 8.0+
* Python 3.12 (Standard Windows installation)
* Dependencies: `pythonnet`, `matplotlib`

## ⚙️ Installation & Setup
1. Clone the repository:
   ```bash
   git clone https://github.com/Atharva-Ramawat/DWSIM-python-automation.git
