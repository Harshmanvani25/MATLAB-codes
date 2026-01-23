# -*- coding: utf-8 -*-
"""
Plasma Diagnostics Data Trimming
Plasma start/end detected from IP7
Time-based trimming applied to all diagnostics
Sampling frequency is known and enforced
Author: Harsh Manvani
"""

# ============================================================
# IMPORTS
# ============================================================

import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os

plt.close('all')

# ============================================================
# USER CONFIGURATION
# ============================================================

EXCEL_FILE = "32322.xlsx"

SHEET_IP7     = "IP7"
SHEET_HALPHA  = "HAlpha"
SHEET_MIRNOV1 = "MIRNOV1"

INPUT_BASE_PATH = r"D:\IPR SEM 8 Internship\Datasets\Halpha_shots_dataset_experimentation\DATA_set for ML"
input_file_path = os.path.join(INPUT_BASE_PATH, EXCEL_FILE)

OUTPUT_BASE_PATH = r"D:\IPR SEM 8 Internship\codes\Perfect code"
os.makedirs(OUTPUT_BASE_PATH, exist_ok=True)

output_file = EXCEL_FILE.replace(".xlsx", "_new.xlsx")
output_file_path = os.path.join(OUTPUT_BASE_PATH, output_file)

# Known sampling frequency (ADC truth)
FS = 100_0000        # Hz  <<< CHANGE IF REQUIRED

# Plasma end threshold (IP7 based)
NEGATIVE_THRESHOLD = -3

# ============================================================
# LOAD DATA
# ============================================================

ip7_df     = pd.read_excel(input_file_path, sheet_name=SHEET_IP7, header=None)
halpha_df  = pd.read_excel(input_file_path, sheet_name=SHEET_HALPHA, header=None)
mirnov_df  = pd.read_excel(input_file_path, sheet_name=SHEET_MIRNOV1, header=None)

for df in [ip7_df, halpha_df, mirnov_df]:
    df.columns = ["Time_ms", "Signal"]
    df.sort_values("Time_ms", inplace=True)
    df.reset_index(drop=True, inplace=True)

# ============================================================
# DETECT PLASMA START (FROM IP7)
# ============================================================

if 0.0 not in ip7_df["Time_ms"].values:
    raise ValueError("Exact 0 ms not found in IP7 time column")

start_time_ms = 0.0

# ============================================================
# DETECT PLASMA END (FROM IP7 THRESHOLD)
# ============================================================

end_time_ms = None
start_idx = ip7_df.index[ip7_df["Time_ms"] == start_time_ms][0]

for i in range(start_idx + 1, len(ip7_df)):
    if ip7_df.loc[i, "Signal"] <= NEGATIVE_THRESHOLD:
        end_time_ms = ip7_df.loc[i, "Time_ms"]
        break

if end_time_ms is None:
    end_time_ms = ip7_df["Time_ms"].iloc[-1]

# ============================================================
# TIME-BASED TRIMMING FUNCTION (SAFE)
# ============================================================

def trim_by_time(df, start_ms, end_ms, fs):
    """
    Trim signal using time (NOT index)
    Reconstructs perfect time axis from fs
    """
    out = df[
        (df["Time_ms"] >= start_ms) &
        (df["Time_ms"] <= end_ms)
    ].copy()

    out.reset_index(drop=True, inplace=True)

    # Rebuild DSP-accurate time axis
    n = len(out)
    out["Time_sec"] = np.arange(n) / fs
    out["Time_ms_fs"] = out["Time_sec"] * 1e3

    return out

# ============================================================
# APPLY TRIMMING TO ALL SIGNALS
# ============================================================

ip7_clean     = trim_by_time(ip7_df, start_time_ms, end_time_ms, FS)
halpha_clean  = trim_by_time(halpha_df, start_time_ms, end_time_ms, FS)
mirnov_clean  = trim_by_time(mirnov_df, start_time_ms, end_time_ms, FS)

# ============================================================
# SAVE CLEAN DATA TO NEW EXCEL FILE
# ============================================================

with pd.ExcelWriter(output_file_path) as writer:
    ip7_clean.to_excel(writer, sheet_name="IP7", index=False)
    halpha_clean.to_excel(writer, sheet_name="HAlpha", index=False)
    mirnov_clean.to_excel(writer, sheet_name="MIRNOV1", index=False)

# ============================================================
# PLOTTING – FULL SIGNALS (ORIGINAL TIME)
# ============================================================

plt.figure(figsize=(14, 9))
plt.suptitle("Full Signals with Detected Plasma Start & End", fontsize=14)

plt.subplot(3, 1, 1)
plt.plot(ip7_df["Time_ms"], ip7_df["Signal"], label="IP7")
plt.axvline(start_time_ms, color='g', linestyle='--', label="Start (0 ms)")
plt.axvline(end_time_ms, color='r', linestyle='--', label="End")
plt.axhline(NEGATIVE_THRESHOLD, color='k', linestyle=':', label="Threshold")
plt.ylabel("IP7")
plt.legend()
plt.grid(True)

plt.subplot(3, 1, 2)
plt.plot(halpha_df["Time_ms"], halpha_df["Signal"], label="HAlpha")
plt.axvline(start_time_ms, color='g', linestyle='--')
plt.axvline(end_time_ms, color='r', linestyle='--')
plt.ylabel("HAlpha")
plt.legend()
plt.grid(True)

plt.subplot(3, 1, 3)
plt.plot(mirnov_df["Time_ms"], mirnov_df["Signal"], label="Mirnov1")
plt.axvline(start_time_ms, color='g', linestyle='--')
plt.axvline(end_time_ms, color='r', linestyle='--')
plt.xlabel("Time (ms)")
plt.ylabel("Mirnov")
plt.legend()
plt.grid(True)

plt.tight_layout(rect=[0, 0, 1, 0.95])
plt.show()

# ============================================================
# PLOTTING – PLASMA-ONLY SIGNALS (FS-BASED TIME)
# ============================================================

plt.figure(figsize=(14, 9))
plt.suptitle("Plasma-Only Signals (Time from Sampling Frequency)", fontsize=14)

plt.subplot(3, 1, 1)
plt.plot(ip7_clean["Time_ms_fs"], ip7_clean["Signal"])
plt.ylabel("IP7")
plt.grid(True)

plt.subplot(3, 1, 2)
plt.plot(halpha_clean["Time_ms_fs"], halpha_clean["Signal"])
plt.ylabel("HAlpha")
plt.grid(True)

plt.subplot(3, 1, 3)
plt.plot(mirnov_clean["Time_ms_fs"], mirnov_clean["Signal"])
plt.xlabel("Time (ms)")
plt.ylabel("Mirnov")
plt.grid(True)

plt.tight_layout(rect=[0, 0, 1, 0.95])
plt.show()

# ============================================================
# FINAL SUMMARY
# ============================================================

print("Processing completed successfully")
print(f"Output file saved at:\n{output_file_path}")
print(f"Sampling frequency : {FS:.2f} Hz")
print(f"Nyquist frequency  : {FS/2:.2f} Hz")
print(f"Plasma duration    : {len(ip7_clean)/FS*1e3:.3f} ms")
print(f"IP7 samples kept   : {len(ip7_clean)}")
