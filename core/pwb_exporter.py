# core/pwb_exporter.py

import os
import win32com.client


def export_violation_ctg(pwb_path: str, log_func) -> str:
    """
    Core logic that talks to PowerWorld SimAuto and exports
    the ViolationCTG table to CSV.

    Returns:
        Path to the CSV file that was written.

    Raises:
        RuntimeError on PowerWorld/SimAuto errors.
    """

    base, _ = os.path.splitext(pwb_path)
    csv_out = base + "_ViolationCTG.csv"

    log_func("Connecting to PowerWorld via SimAuto...")
    simauto = win32com.client.Dispatch("pwrworld.SimulatorAuto")
    log_func("Connected.")

    try:
        # 1) Open the case (must already have contingency results stored)
        log_func(f"Opening case: {pwb_path}")
        (err,) = simauto.OpenCase(pwb_path)
        if err:
            raise RuntimeError(f"OpenCase error: {err}")
        log_func("Case opened successfully; using existing contingency results.")

        # 2) Enter Contingency mode so ViolationCTG is active
        log_func("Entering Contingency mode...")
        (err,) = simauto.RunScriptCommand("EnterMode(Contingency);")
        if err:
            raise RuntimeError(f"EnterMode(Contingency) error: {err}")

        # 3) Save ViolationCTG table to CSV
        log_func(f"Saving ViolationCTG data to CSV:\n  {csv_out}")
        clean_csv = csv_out.replace("\\", "/")  # avoid backslash issues in script
        cmd = (
            f'SaveData("{clean_csv}", CSV, ViolationCTG, '
            "[ALL], [], \"\");"
        )
        (err,) = simauto.RunScriptCommand(cmd)
        if err:
            raise RuntimeError(f"SaveData(ViolationCTG) error: {err}")
        log_func("CSV export complete for ViolationCTG.")

    finally:
        # Clean up SimAuto
        try:
            simauto.CloseCase()
        except Exception:
            pass
        del simauto

    return csv_out
