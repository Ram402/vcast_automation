"""
VectorCAST Batch Compilation Runner
====================================
Compiles every module listed in MODULES using the exact same logic as
vcast_auto_compile3.py — without modifying that script at all.

How it works
------------
1.  For each module the runner searches HEADER_SEARCH_ROOT for a .c file
    whose stem matches the UUT name.
2.  It derives SOURCE_DIR from the directory that contains the found .c file.
3.  It overrides all the global variables in vcast_auto_compile3 to point at
    the current module, then calls compile_one_module() which runs the full
    auto-retry loop (header fix + macro fix + linker stub).
4.  If a module fails for any reason the exception is caught, the failure is
    recorded, and compilation continues for the next module.
5.  A final summary table is printed showing PASS / FAIL for every module.

Configuration — edit the section below
---------------------------------------
"""

import os
import sys
import importlib.util
import traceback
from datetime import datetime
from pathlib import Path

# ============================================================================
# BATCH CONFIGURATION  –  edit these
# ============================================================================

VECTORCAST_DIR     = r"C:\VCAST"
BASE_DIR_NAME      = "R2"
BASE_DIR_PATH      = r"D:\project_4\BC4i_P E2.0_B2412\B2412"
HEADER_SEARCH_ROOT = BASE_DIR_PATH
WORKSPACE_ROOT     = r"D:\Nagesh\workspace\UT"   # each module gets its own subfolder here
MAX_RETRY_ROUNDS   = 100

# The original compilation script (keep it next to this file, or give full path)
COMPILE_SCRIPT = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                              "vcast_auto_compile3.py")

# -----------------------------------------------------------------------
# MODULES TABLE
# Each entry:  (UUT_stem, .c_filename)
#   UUT_stem   – used for ENV_NAME, WORK_DIR subfolder, ENVIRO.UUT
#   .c_filename – the actual source file to find under HEADER_SEARCH_ROOT
# -----------------------------------------------------------------------
MODULES = [

    ("Aswc_Brake_Lp",                       "Aswc_Brake_Lp.c"),
    ("PDC_BrakeLamp",                  "PDC_BrakeLamp.c"),

]
# ============================================================================
# BATCH LOG
# ============================================================================
BATCH_LOG = os.path.join(WORKSPACE_ROOT, "batch_compile_log.txt")


def blog(msg: str, also_print: bool = True) -> None:
    os.makedirs(WORKSPACE_ROOT, exist_ok=True)
    with open(BATCH_LOG, "a", encoding="utf-8") as f:
        f.write(msg + "\n")
    if also_print:
        print(msg)


# ============================================================================
# FIND .c FILE ON DISK
# ============================================================================

def find_c_file(filename: str, root: str) -> str:
    """
    Case-insensitive walk of root to find filename.
    Returns the full path of the first match, or '' if not found.
    """
    target = filename.lower()
    for dirpath, _dirs, files in os.walk(root):
        for fn in files:
            if fn.lower() == target:
                return os.path.join(dirpath, fn)
    return ""


# ============================================================================
# LOAD THE ORIGINAL COMPILE SCRIPT AS A MODULE
# ============================================================================

def _load_compile_module():
    """
    Dynamically load vcast_auto_compile3.py as a Python module object.
    We load it once and re-use it, overriding its globals per module.
    """
    spec = importlib.util.spec_from_file_location("vcast_compile", COMPILE_SCRIPT)
    mod  = importlib.util.module_from_spec(spec)
    # Execute the module body (defines all functions, sets initial globals)
    spec.loader.exec_module(mod)
    return mod


# ============================================================================
# COMPILE ONE MODULE
# ============================================================================

def compile_one_module(mod, uut_stem: str, c_file: str, source_dir: str) -> bool:
    """
    Override the globals in *mod* to point at this module, then run main().

    Returns True on success, False on failure.
    """
    work_dir   = os.path.join(WORKSPACE_ROOT, uut_stem)
    env_name   = uut_stem
    env_script = os.path.join(work_dir, f"{env_name}.env")
    cfg_file   = os.path.join(work_dir, "CCAST_.CFG")
    build_log  = os.path.join(work_dir, "build_log.txt")
    detailed   = os.path.join(work_dir, "detailed_log.txt")
    error_log  = os.path.join(work_dir, "error_log.txt")

    os.makedirs(work_dir, exist_ok=True)

    # ── Patch ALL relevant globals in the compile module ──────────────
    mod.ENV_NAME           = env_name
    mod.WORK_DIR           = work_dir
    mod.UUT_FILE           = uut_stem
    mod.SOURCE_DIR_1       = source_dir
    mod.SOURCE_DIR_2       = ""
    mod.SOURCE_DIR_3       = ""
    mod.BASE_DIR_NAME      = BASE_DIR_NAME
    mod.BASE_DIR_PATH      = BASE_DIR_PATH
    mod.VECTORCAST_DIR     = VECTORCAST_DIR
    mod.HEADER_SEARCH_ROOT = HEADER_SEARCH_ROOT
    mod.MAX_RETRY_ROUNDS   = MAX_RETRY_ROUNDS
    mod.DEFINES            = ["__USE_MINGW_ANSI_STDIO"]   # reset per module
    mod.EXTRA_INCLUDE_1    = ""
    mod.EXTRA_INCLUDE_2    = ""
    mod.EXTRA_INCLUDE_3    = ""

    # Derived paths
    mod.BUILD_LOG    = build_log
    mod.DETAILED_LOG = detailed
    mod.ERROR_LOG    = error_log
    mod.CLICAST_EXE  = os.path.join(VECTORCAST_DIR, "clicast.exe")
    mod.ENV_SCRIPT   = env_script
    mod.CFG_FILE     = cfg_file

    # Override fail() so it raises instead of calling sys.exit()
    # (allows the batch runner to catch the failure and move on)
    def _fail_raise(reason: str):
        raise RuntimeError(f"[FAILED] {reason}")
    mod.fail = _fail_raise

    # Override show_alert() to suppress popups during batch run
    mod.show_alert = lambda title, message, icon="Information": None

    # Override input() prompt so it doesn't block during batch
    import builtins
    _orig_input = builtins.input
    builtins.input = lambda prompt="": None

    try:
        mod.main()
        return True
    except (RuntimeError, SystemExit) as exc:
        blog(f"  [FAIL] {uut_stem}: {exc}")
        return False
    except Exception as exc:
        blog(f"  [FAIL] {uut_stem}: Unexpected error: {exc}")
        blog(traceback.format_exc(), also_print=False)
        return False
    finally:
        builtins.input = _orig_input


# ============================================================================
# MAIN BATCH RUNNER
# ============================================================================

def main():
    sep  = "=" * 76
    sep2 = "─" * 76

    # Init batch log
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    os.makedirs(WORKSPACE_ROOT, exist_ok=True)
    with open(BATCH_LOG, "w", encoding="utf-8") as f:
        f.write(f"{sep}\nVectorCAST Batch Compilation\nStarted: {now}\n{sep}\n\n")
        f.write(f"  Compile script : {COMPILE_SCRIPT}\n")
        f.write(f"  BASE_DIR_PATH  : {BASE_DIR_PATH}\n")
        f.write(f"  WORKSPACE_ROOT : {WORKSPACE_ROOT}\n")
        f.write(f"  Modules        : {len(MODULES)}\n\n")

    print(sep)
    print("  VectorCAST Batch Compilation Runner")
    print(f"  Modules       : {len(MODULES)}")
    print(f"  Base dir      : {BASE_DIR_PATH}")
    print(f"  Workspace     : {WORKSPACE_ROOT}")
    print(f"  Compile script: {COMPILE_SCRIPT}")
    print(sep)

    # Verify compile script exists
    if not os.path.isfile(COMPILE_SCRIPT):
        print(f"\n[ERROR] Compile script not found: {COMPILE_SCRIPT}")
        print("Place vcast_auto_compile3.py in the same folder as this script.")
        sys.exit(1)

    # Load the compile module once
    blog("\n[BATCH] Loading compile module...")
    try:
        mod = _load_compile_module()
        blog(f"  [OK] Loaded: {COMPILE_SCRIPT}")
    except Exception as exc:
        blog(f"[ERROR] Failed to load compile script: {exc}")
        sys.exit(1)

    results = []   # list of (uut_stem, c_file, source_dir, status, duration_s)

    for idx, (uut_stem, c_filename) in enumerate(MODULES, 1):

        print()
        print(sep2)
        print(f"  MODULE {idx}/{len(MODULES)}: {uut_stem}")
        print(f"  Source file  : {c_filename}")
        print(sep2)
        blog(f"\n[MODULE {idx}/{len(MODULES)}] {uut_stem}  ({c_filename})")

        # ── Step 1: find the .c file ──────────────────────────────────
        blog(f"  Searching for {c_filename} under {HEADER_SEARCH_ROOT} ...")
        c_path = find_c_file(c_filename, HEADER_SEARCH_ROOT)

        if not c_path:
            msg = f"Source file '{c_filename}' not found under '{HEADER_SEARCH_ROOT}'"
            blog(f"  [SKIP] {msg}")
            print(f"\n  [SKIP] {msg}")
            results.append((uut_stem, c_filename, "", "SKIP – file not found", 0))
            continue

        source_dir = os.path.dirname(c_path)
        blog(f"  [FOUND] {c_path}")
        blog(f"  Source dir   : {source_dir}")
        print(f"  Found        : {c_path}")
        print(f"  Source dir   : {source_dir}")

        # ── Step 2: compile ───────────────────────────────────────────
        t_start = datetime.now()
        print(f"\n  Starting compilation at {t_start.strftime('%H:%M:%S')} ...")
        blog(f"  Compiling {uut_stem}...")

        success = compile_one_module(mod, uut_stem, c_filename, source_dir)

        t_end  = datetime.now()
        elapsed = round((t_end - t_start).total_seconds())
        status  = "PASS" if success else "FAIL"

        blog(f"  Result: {status}  ({elapsed}s)")
        print(f"\n  Result: {status}  (elapsed: {elapsed}s)")

        results.append((uut_stem, c_filename, source_dir, status, elapsed))

    # ================================================================
    #  FINAL SUMMARY TABLE
    # ================================================================
    print()
    print(sep)
    print("  BATCH COMPILATION SUMMARY")
    print(sep)
    print(f"  {'MODULE':<35} {'FILE':<35} {'STATUS':<8} {'TIME':>6}")
    print(f"  {'─'*35} {'─'*35} {'─'*8} {'─'*6}")

    passed = failed = skipped = 0
    for uut, cfile, sdir, status, elapsed in results:
        icon  = "✓" if status == "PASS" else ("~" if "SKIP" in status else "✗")
        tstr  = f"{elapsed}s" if elapsed else "-"
        print(f"  {icon} {uut:<34} {cfile:<35} {status:<8} {tstr:>6}")
        if status == "PASS":
            passed += 1
        elif "SKIP" in status:
            skipped += 1
        else:
            failed += 1

    print(f"  {'─'*35} {'─'*35} {'─'*8} {'─'*6}")
    print(f"  Total: {len(results)}   PASS: {passed}   FAIL: {failed}   SKIP: {skipped}")
    print(sep)

    # Write summary to batch log
    completed = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    blog(f"\n{sep}")
    blog("BATCH SUMMARY")
    blog(sep)
    for uut, cfile, sdir, status, elapsed in results:
        tstr = f"{elapsed}s" if elapsed else "-"
        blog(f"  {status:<6} {uut:<35} {cfile}  [{tstr}]")
        if sdir:
            blog(f"         Source: {sdir}", also_print=False)
    blog(f"\nCompleted: {completed}")
    blog(f"PASS: {passed}  FAIL: {failed}  SKIP: {skipped}")
    blog(sep)

    print(f"\n  Batch log : {BATCH_LOG}")
    print(f"  Each module has its own detailed_log.txt in:")
    print(f"    {WORKSPACE_ROOT}\\<MODULE_NAME>\\")

    if failed > 0:
        print(f"\n  [NOTE] {failed} module(s) failed.")
        print("  Check <MODULE_NAME>\\build_log.txt for the specific error.")

    print("\nDone.")
    os.system("exit")

if __name__ == "__main__":
    main()
