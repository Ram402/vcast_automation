"""
vcast_it_compile.py (CORRECTED - BACKWARD TRACING)
==================================================
VectorCAST Integration Test (IT) Compilation Script
  - BACKWARD call-graph traversal (who calls UUT functions)
  - BFS-based multi-UUT tracing with RTE stop-boundary
  - Automatic SBF list generation (interface files that call UUT)
  - Graphviz .dot + CSV + Excel dependency reports
  - Reuses vcast_auto_compile3.py for ALL build/fix/coverage logic
    (zero duplication of header-fix, macro-fix, retry, coverage)

DESIGN RULES
------------
  1. vcast_auto_compile3.py is NEVER modified.
  2. All build mechanics are delegated to it via importlib global injection
     (same technique as vcast_batch_compile.py).
  3. IT_UUTS  → ENVIRO.UUT  entries
     SBF_FILES → ENVIRO.SBF  entries (interface files calling UUT)
     A file NEVER appears in both lists.
  4. Dependency discovery does ONE filesystem walk and builds three indexes;
     all subsequent lookups are O(1) dict accesses.
  5. BACKWARD TRACING: For each UUT function, find who calls it (interface).
"""

# ── stdlib ────────────────────────────────────────────────────────────────────
import os
import sys
import re
import csv
import importlib.util
import traceback
import builtins
import shutil
from collections import deque
from datetime import datetime
from pathlib import Path

# ── optional: openpyxl ────────────────────────────────────────────────────────
try:
    import openpyxl
    from openpyxl.styles import (PatternFill, Font, Alignment,
                                  Border, Side, GradientFill)
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# =============================================================================
# ██████████████████████████████  CONFIGURATION  ██████████████████████████████
# =============================================================================

VECTORCAST_DIR = r"C:\VCAST"

BASE_DIR_NAME  = "R"
BASE_DIR_PATH  = r"D:\project_4\BC4i_P E2.0_B2412\B2412"

# Root used for both header search AND dependency scan
HEADER_SEARCH_ROOT = BASE_DIR_PATH

# Where the IT environment and all output artefacts are written
WORK_DIR = r"D:\Workspace_IT\PDC_LampGroup_IT"

# VectorCAST environment name
ENV_NAME = "PDC_LampGroup_IT"

# How many retry rounds to allow the auto-fix engine
MAX_RETRY_ROUNDS = 100

# ── UUT files for this IT environment ────────────────────────────────────────
IT_UUTS = [
    "Aswc_Lugg_Cargo_Lp.c",
    "PDC_LuggageorCargoLamp.c"
]

IT_SBFS = [
    "Asw_Init.c",
    "App_Mode.c",
    "Rte.c"
]

# Additional compiler intrinsic stubs written as ENVIRO.ADDITIONAL_STUB
# (matches the working manual .env which has __DI and __EI)
ADDITIONAL_STUBS: list = ["__DI", "__EI"]

# ── Compiler defines (baseline; auto-extended by macro-fix engine) ────────────
DEFINES: list = ["__USE_MINGW_ANSI_STDIO"]

# ── Extra include seed paths (auto-extended by header-fix engine) ─────────────
EXTRA_INCLUDE_1 = ""
EXTRA_INCLUDE_2 = ""
EXTRA_INCLUDE_3 = ""

# ── Compile script path ───────────────────────────────────────────────────────
COMPILE_SCRIPT = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "vcast_auto_compile3.py"
)

# =============================================================================
# ████████████████████████████  DERIVED PATHS  ████████████████████████████████
# =============================================================================


DETAILED_LOG  = os.path.join(WORK_DIR, "detailed_log.txt")
BUILD_LOG     = os.path.join(WORK_DIR, "build_log.txt")
ERROR_LOG     = os.path.join(WORK_DIR, "error_log.txt")
ENV_SCRIPT    = os.path.join(WORK_DIR, f"{ENV_NAME}.env")
CFG_FILE      = os.path.join(WORK_DIR, "CCAST_.CFG")
CLICAST_EXE   = os.path.join(VECTORCAST_DIR, "clicast.exe")


# =============================================================================
# ████████████████████  OUTPUT GENERATORS  ████████████████████████████████████
# =============================================================================

def _stem(path: str) -> str:
    return os.path.splitext(os.path.basename(path))[0]



# =============================================================================
# ████████████  ENV SCRIPT WRITER (IT version)  ██████████████████████████████
# =============================================================================

def write_it_env_script(
    uut_stems:      list,
    sbf_stems:      list,
    extra_includes: list,
    source_dirs:    list,
    additional_stubs: list = None,
) -> None:
    if additional_stubs is None:
        additional_stubs = ADDITIONAL_STUBS
    """
    Write ENV_NAME.env with:
      ENVIRO.UUT  for every UUT stem
      ENVIRO.SBF  for every SBF stem
      ENVIRO.SEARCH_LIST for source_dirs + extra_includes
    """
    lines = [
        "-- VectorCAST IT Environment Script (auto-generated by vcast_it_compile.py)",
        f"-- Environment : {ENV_NAME}",
        "--",
        "ENVIRO.NEW",
        f"ENVIRO.NAME: {ENV_NAME}",
        "ENVIRO.COVERAGE_TYPE: Statement+MCDC",
        "ENVIRO.WHITE_BOX: YES",
        "ENVIRO.STUB: ALL_BY_PROTOTYPE",
        "ENVIRO.COMPILER: CC",
        "ENVIRO.TYPE_HANDLED_DIRS_ALLOWED:",
        f"ENVIRO.BASE_DIRECTORY: {BASE_DIR_NAME}={BASE_DIR_PATH}",
    ]

    # UUT entries  (interleaved with SBF before UUT, matching manual .env order)
    for stem in sbf_stems:
        lines.append(f"ENVIRO.STUB_BY_FUNCTION: {stem}")

    for stem in uut_stems:
        lines.append(f"ENVIRO.UUT: {stem}")

    # Additional stubs for compiler intrinsics (from working manual .env)
    for stub in (additional_stubs or []):
        lines.append(f"ENVIRO.ADDITIONAL_STUB: {stub}")

    # Search list entries
    norm_base = os.path.normcase(os.path.normpath(BASE_DIR_PATH))

    def _to_env_path(abs_path: str) -> str:
        """Convert absolute path to $(BASE_DIR_NAME)\relative form matching manual .env"""
        try:
            rel = os.path.normpath(os.path.relpath(abs_path, BASE_DIR_PATH))
            return f"$({BASE_DIR_NAME})\\{rel}"
        except ValueError:
            return abs_path

    for sd in source_dirs:
        if sd and os.path.isdir(sd):
            lines.append(f"ENVIRO.SEARCH_LIST: {_to_env_path(sd)}")

    for inc in extra_includes:
        if inc and os.path.isdir(inc):
            lines.append(f"ENVIRO.SEARCH_LIST: {_to_env_path(inc)}")

    # User globals (unchanged from auto_compile3 pattern)
    user_globals = (
        "ENVIRO.USER_GLOBALS:\n"
        "/*************************************************************\n"
        " S0000008.c – variable definitions for user code.\n"
        "*************************************************************/\n"
        "#ifndef VCAST_USER_GLOBALS_EXTERN\n"
        "#define VCAST_USER_GLOBALS_EXTERN\n"
        "#endif\n"
        "#ifdef __cplusplus\n"
        'extern "C"{\n'
        "#endif\n"
        "  VCAST_USER_GLOBALS_EXTERN int  VECTORCAST_INT1;\n"
        "  VCAST_USER_GLOBALS_EXTERN int  VECTORCAST_INT2;\n"
        "  VCAST_USER_GLOBALS_EXTERN int  VECTORCAST_INT3;\n"
        "#ifndef VCAST_NO_FLOAT\n"
        "  VCAST_USER_GLOBALS_EXTERN float VECTORCAST_FLT1;\n"
        "#endif\n"
        "  VCAST_USER_GLOBALS_EXTERN char VECTORCAST_STR1[8];\n"
        "  VCAST_USER_GLOBALS_EXTERN int  VECTORCAST_BUFFER[4];\n"
        "#ifdef __cplusplus\n"
        "}\n"
        "#endif\n"
        "ENVIRO.END_USER_GLOBALS:\n"
        "ENVIRO.END\n"
    )

    with open(ENV_SCRIPT, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n" + user_globals)



# =============================================================================
# ████████████████████  CCAST_.CFG WRITER  ████████████████████████████████████
# =============================================================================

def write_it_cfg(defines: list) -> None:
    cfg_lines = [
        "C_COMPILER_CFG_SOURCE: PY_CONFIGURATOR",
        "C_COMPILER_FAMILY_NAME: GNU_Native",
        "C_COMPILER_HIERARCHY_STRING: VectorCAST MinGW_C",
        "C_COMPILER_OUTPUT_FLAG: -o",
        "C_COMPILER_PY_ARGS: --lang c --version Built-in-MinGW",
        "C_COMPILER_TAG: BUILTIN_MINGW_C",
        "C_COMPILER_VERSION_CMD: gcc --version",
        "C_COMPILE_CMD: gcc -c -g",
        "C_DEBUG_CMD: gdb",
        "C_EDG_FLAGS: -w --gcc --gnu_version 100200 --64_bit_target --x86_64 --mingw",
        "C_LINKER_VERSION_CMD: ld --version",
        "C_LINK_CMD: gcc -g",
        "C_PREPROCESS_CMD: gcc -E -ftrack-macro-expansion=0 -C",
        "VARIANT_LOGICS_PATH: ",
        "VCAST_ASSEMBLY_FILE_EXTENSIONS: s",
        "VCAST_COLLAPSE_STD_HEADERS: COLLAPSE_NONE",
        "VCAST_COMMAND_LINE_DEBUGGER: TRUE",
        "VCAST_DISABLE_STD_WSTRING_DETECTION: TRUE",
        "VCAST_DISPLAY_UNINST_EXPR: FALSE",
        "VCAST_ENVIRONMENT_FILES: ",
        "VCAST_GNU_SYSTEM_MARKER: TRUE",
        "VCAST_HAS_LONGLONG: TRUE",
        f"VCAST_PREPEND_TO_PATH_DIRS: $(VECTORCAST_DIR)/MinGW/bin",
        "VCAST_TEST_VALUES_DICTIONARY: ",
        "VCAST_TYPEOF_OPERATOR: TRUE",
        "VCAST_VCDB_FLAG_STRING: -isystem=1",
        "VCDB_CMD_VERB: ",
        "VCDB_FILENAME: ",
        "WHITEBOX: YES",
        # ── Coverage instrumentation keys (from working CCAST_.CFG) ──────────
        "VCAST_COVERAGE_SOURCE_FILE_PERSPECTIVE: FALSE",
        "VCAST_DISPLAY_FUNCTION_COVERAGE: TRUE",
        "VCAST_ENABLE_FUNCTION_AND_CALL_COVERAGE: TRUE",
        "VCAST_ENABLE_FUNCTION_CALL_COVERAGE: TRUE",
    ]
    with open(CFG_FILE, "w", encoding="utf-8") as f:
        f.write("\n".join(cfg_lines) + "\n")
        # FIX: VectorCAST requires one C_DEFINE_LIST entry per define.
        # Writing all defines space-separated on one line causes them to be
        # silently ignored, which prevents any auto-discovered guard from applying.
        for d in defines:
            f.write(f"C_DEFINE_LIST: {d}\n")
    #tlog(f"[BUILD] CFG written. Defines ({len(defines)}): {' '.join(defines)}")


# =============================================================================
# ████████████████████  DELEGATE TO vcast_auto_compile3  ██████████████████████
# =============================================================================

def _load_compile_mod():
    if not os.path.isfile(COMPILE_SCRIPT):
        raise FileNotFoundError(f"vcast_auto_compile3.py not found at {COMPILE_SCRIPT}")
    spec = importlib.util.spec_from_file_location("vcast_compile", COMPILE_SCRIPT)
    mod  = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def run_it_build(
    mod,
    uut_stems:      list[str],
    sbf_stems:      list[str],
    source_dirs:    list[str],
    extra_includes: list[str],
    active_defines: list,
) -> bool:
    """
    Patch globals on `mod` (vcast_auto_compile3) for IT build.

    We override write_env_script so it writes an IT-style env (multi-UUT + SBF).
    We override write_cfg so it uses our active_defines list.
    Everything else (retry loop, header fix, macro fix, coverage) runs unmodified.
    """

    env_dir = os.path.join(WORK_DIR, ENV_NAME)

    # ── patch globals ─────────────────────────────────────────────────────────
    mod.ENV_NAME           = ENV_NAME
    mod.WORK_DIR           = WORK_DIR
    mod.UUT_FILE           = uut_stems[0] if uut_stems else ""   # primary UUT
    mod.SOURCE_DIR_1       = source_dirs[0] if len(source_dirs) > 0 else ""
    mod.SOURCE_DIR_2       = source_dirs[1] if len(source_dirs) > 1 else ""
    mod.SOURCE_DIR_3       = source_dirs[2] if len(source_dirs) > 2 else ""
    mod.BASE_DIR_NAME      = BASE_DIR_NAME
    mod.BASE_DIR_PATH      = BASE_DIR_PATH
    mod.VECTORCAST_DIR     = VECTORCAST_DIR
    mod.HEADER_SEARCH_ROOT = HEADER_SEARCH_ROOT
    mod.MAX_RETRY_ROUNDS   = MAX_RETRY_ROUNDS
    mod.DEFINES            = active_defines
    mod.EXTRA_INCLUDE_1    = extra_includes[0] if len(extra_includes) > 0 else ""
    mod.EXTRA_INCLUDE_2    = extra_includes[1] if len(extra_includes) > 1 else ""
    mod.EXTRA_INCLUDE_3    = extra_includes[2] if len(extra_includes) > 2 else ""

    mod.BUILD_LOG    = BUILD_LOG
    mod.DETAILED_LOG = DETAILED_LOG
    mod.ERROR_LOG    = ERROR_LOG
    mod.CLICAST_EXE  = CLICAST_EXE
    mod.ENV_SCRIPT   = ENV_SCRIPT
    mod.CFG_FILE     = CFG_FILE

    # ── override write_env_script to produce IT-style .env ───────────────────
    def _write_env_it(extra_includes_arg: list) -> None:
        write_it_env_script(uut_stems, sbf_stems, extra_includes_arg,
                            source_dirs, ADDITIONAL_STUBS)

    mod.write_env_script = _write_env_it

    # ── override write_cfg to use our shared active_defines list ─────────────
    def _write_cfg_it() -> None:
        write_it_cfg(active_defines)

    mod.write_cfg = _write_cfg_it

    # ── suppress interactive popups / sys.exit ────────────────────────────────
    mod.show_alert = lambda title, message, icon="Information": None

    def _fail_raise(reason: str):
        raise RuntimeError(reason)

    mod.fail = _fail_raise

    _orig_input = builtins.input
    builtins.input = lambda prompt="": None

    try:
        #tlog(f"\n[BUILD] Delegating to vcast_auto_compile3 …")
        result = mod.main()
        #tlog(f"[BUILD] Auto-compile returned: {result}")
        return result if isinstance(result, bool) else True
    except (RuntimeError, SystemExit) as exc:
        #tlog(f"[BUILD] Auto-compile raised: {exc}")
        return False
    except Exception as exc:
        #tlog(f"[BUILD] Auto-compile failed: {exc}")
        #tlog(f"[BUILD] Traceback: {traceback.format_exc()}")
        return False
    finally:
        builtins.input = _orig_input


# =============================================================================
# ████████████████████████████████  MAIN  █████████████████████████████████████
# =============================================================================

def main() -> None:
    sep  = "=" * 76
    sep2 = "─" * 76

    os.makedirs(WORK_DIR, exist_ok=True)

    print(sep)
    print("  VectorCAST IT Compile  |  Backward Dependency Trace + Auto-Build")
    print(f"  Environment : {ENV_NAME}")
    print(f"  UUTs        : {len(IT_UUTS)}")
    for u in IT_UUTS:
        print(f"    • {u}")
    print(f"  Search root : {HEADER_SEARCH_ROOT}")
    print(f"  Work dir    : {WORK_DIR}")
    print(sep)

    # =========================================================================
    # STEP 1 — Verify clicast
    # =========================================================================
    #tlog(f"\n[BUILD] STEP 1 — Checking clicast.exe …")
    if not os.path.isfile(CLICAST_EXE):
        err = f"clicast.exe not found: {CLICAST_EXE}"
        print(f"ERROR: {err}")
        #tlog(f"[BUILD] ERROR: {err}")
        return
    #tlog(f"[BUILD]   OK: {CLICAST_EXE}")

    print()
    print(f"{'─'*76}")
    print(f"  Setting environment …")
    print(f"{'─'*76}")

    mingw_bin = os.path.join(VECTORCAST_DIR, "MinGW", "bin")
    os.environ["PATH"] = f"{VECTORCAST_DIR};{mingw_bin};{os.environ.get('PATH', '')}"
    os.environ.pop("VECTORCAST_DIR", None)
    os.chdir(WORK_DIR)
    
    uut_stems = [
        os.path.splitext(x)[0]
        for x in IT_UUTS
    ]

    sbf_stems = [
        os.path.splitext(x)[0]
        for x in IT_SBFS
    ]

    source_dirs = []

    def add_dir_if_missing(path):
        if path and os.path.isdir(path) and path not in source_dirs:
            source_dirs.append(path)

    #
    # Find UUT directories
    #
    for uut_fname in IT_UUTS:

        for root, _, files in os.walk(HEADER_SEARCH_ROOT):

            if uut_fname in files:

                add_dir_if_missing(root)
                break

    #
    # Find SBF directories
    #
    for sbf_fname in IT_SBFS:

        for root, _, files in os.walk(HEADER_SEARCH_ROOT):

            if sbf_fname in files:

                add_dir_if_missing(root)
                break
    # =========================================================================
    # STEP 3 — Resolve UUT abs paths
    # =========================================================================
    #tlog(f"\n[TRACE] STEP 3 — Resolving UUT file paths …")
    uut_abs_paths = []
    uut_stems = []
    missing_uuts = []

    for uut_fname in IT_UUTS:

        found_path = None

        for root, _, files in os.walk(HEADER_SEARCH_ROOT):

            if uut_fname in files:
                found_path = os.path.join(root, uut_fname)
                break

        if found_path:
            uut_abs_paths.append(found_path)
            uut_stems.append(os.path.splitext(uut_fname)[0])

            src_dir = os.path.dirname(found_path)

            if src_dir not in source_dirs:
                source_dirs.append(src_dir)

        else:
            missing_uuts.append(uut_fname)
            
    sbf_abs_paths = []
    sbf_stems = []
    missing_sbfs = []

    for sbf_fname in IT_SBFS:

        found_path = None

        for root, _, files in os.walk(HEADER_SEARCH_ROOT):

            if sbf_fname in files:
                found_path = os.path.join(root, sbf_fname)
                break

        if found_path:
            sbf_abs_paths.append(found_path)
            sbf_stems.append(os.path.splitext(sbf_fname)[0])
            
            sbf_dir = os.path.dirname(found_path)

            if sbf_dir not in source_dirs:
                source_dirs.append(sbf_dir)
        else:
            missing_sbfs.append(sbf_fname)

    if missing_uuts:
        print(f"[WARN] Some UUT files not found on disk: {missing_uuts}")
        print("       They will still be listed as ENVIRO.UUT — clicast may find them via SEARCH_LIST.")

    if missing_sbfs:
        print(f"[WARN] Some SBF files not found on disk: {missing_sbfs}")
        print("       They will still be listed as ENVIRO.STUB_BY_FUNCTION.")



    #print()
    print("  UUT list:")
    for s in uut_stems:
        print(f"    • {s}")
    print("  SBF list:")
    for s in sbf_stems:
        print(f"    • {s}")


    mod = _load_compile_mod()
    #tlog(f"[BUILD]   Loaded: {COMPILE_SCRIPT}")

    # Prepare initial defines list (shared mutable — auto-fix engine extends it)
    active_defines = list(DEFINES)

    # Seed includes
    extra_includes = [e for e in [EXTRA_INCLUDE_1, EXTRA_INCLUDE_2, EXTRA_INCLUDE_3] if e]

    
    print("\nSEARCH LIST DIRECTORIES:")
    for d in source_dirs:
        print("  ", d)

    success = run_it_build(
        mod,
        uut_stems,
        sbf_stems,
        source_dirs,
        extra_includes,
        active_defines,
    )

    # =========================================================================
    # FINAL SUMMARY
    # =========================================================================
    print()
    print(sep)
    print("  IT BUILD SUMMARY")
    print(sep)
    print(f"  Environment  : {ENV_NAME}")
    print(f"  Status       : {'SUCCESS ✓' if success else 'FAILED ✗'}")
    print(f"  UUTs ({len(uut_stems)}):")
    for s in uut_stems:
        print(f"    • {s}")
    print(f"  SBFs ({len(sbf_stems)}):")
    for s in sbf_stems:
        print(f"    • {s}")
    print(sep)



if __name__ == "__main__":
    main()