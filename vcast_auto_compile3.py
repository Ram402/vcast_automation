"""
VectorCAST Automated Compilation Script
  - AUTO MISSING-HEADER RESOLUTION
  - AUTO MACRO EXPANSION ERROR BACKTRACING & FIX  (P2VAR / FUNC / VAR / P2CONST etc.)

HOW THE MACRO FIX WORKS
========================
AUTOSAR/MCAL headers define macros like P2VAR, FUNC, VAR twice:

    #if defined(MCAL_SUPPORT_MOBILGENE_1_0) || defined(MCAL_SUPPORT_MOBILGENE_2_0)
    #define P2VAR(ptrtype, memclass, ptrclass)   ptrtype *          <-- simple, safe
    #else
    #define P2VAR(ptrtype, memclass, ptrclass)   ptrclass ptrtype * memclass  <-- breaks MinGW
    #endif

VectorCAST uses MinGW which has neither guard defined → falls into #else → broken expansion.

Fix:
  1. Detect "expected X before Y" parse errors in the log.
  2. Read the offending source line, extract every CAPS_MACRO( call.
  3. Find ALL #define occurrences of that macro in the project headers.
  4. Identify which one is inside an #else block and which #if guards it.
  5. Add that guard macro as a C_DEFINE_LIST entry in CCAST_.CFG.
  6. Rebuild.
"""

import os, sys, re, shutil, subprocess
from datetime import datetime

# ============================================================================
# CONFIGURATION  – edit these values for each project
# ============================================================================

VECTORCAST_DIR = r"C:\VCAST"
ENV_NAME       = "PDC_OptionProcessingConfirm"
WORK_DIR       = r"D:\Rama\Workspace_UT\PDC_OptionProcessingConfirm"

BASE_DIR_NAME  = "R"
BASE_DIR_PATH  = r"D:\project_4\BC4i_P E2.0_B2412\B2412"

SOURCE_DIR_1   = rf"{BASE_DIR_PATH}\Static_Code\KSC\SYSTEMS\Interface\Option_Processing_Confirm"
SOURCE_DIR_2   = ""
SOURCE_DIR_3   = ""

UUT_FILE       = "PDC_OptionProcessingConfirm"

# Extended at runtime by macro-fix engine – do NOT hardcode MCAL guards here,
# the script discovers and adds them automatically.
DEFINES: list = ["__USE_MINGW_ANSI_STDIO"]

EXTRA_INCLUDE_1 = ""
EXTRA_INCLUDE_2 = ""
EXTRA_INCLUDE_3 = ""

HEADER_SEARCH_ROOT = BASE_DIR_PATH
MAX_RETRY_ROUNDS   = 100

# ============================================================================
# DERIVED PATHS  (do not edit)
# ============================================================================
BUILD_LOG    = os.path.join(WORK_DIR, "build_log.txt")
DETAILED_LOG = os.path.join(WORK_DIR, "detailed_log.txt")
ERROR_LOG    = os.path.join(WORK_DIR, "error_log.txt")
CLICAST_EXE  = os.path.join(VECTORCAST_DIR, "clicast.exe")
ENV_SCRIPT   = os.path.join(WORK_DIR, f"{ENV_NAME}.env")
CFG_FILE     = os.path.join(WORK_DIR, "CCAST_.CFG")

# ============================================================================
# LOGGING
# ============================================================================

def log(msg: str, also_print: bool = True) -> None:
    with open(DETAILED_LOG, "a", encoding="utf-8") as f:
        f.write(msg + "\n")
    if also_print:
        print(msg)


def show_alert(title: str, message: str, icon: str = "Information") -> None:
    ps = (
        "Add-Type -AssemblyName PresentationFramework; "
        f"[System.Windows.MessageBox]::Show('{message}', '{title}', 'OK', '{icon}')"
    )
    subprocess.run(["powershell", "-command", ps],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def fail(reason: str) -> None:
    log(f"\n[FAILED] {reason}")
    print(f"\nLog files:\n  {BUILD_LOG}\n  {ERROR_LOG}\n  {DETAILED_LOG}")
    show_alert("VectorCAST Build Failed",
               f"FAILED: {ENV_NAME}\\n{reason}\\nLogs: {WORK_DIR}", "Error")
    input("\nPress Enter to exit...")
    sys.exit(1)


# ============================================================================
# SHARED UTILITY
# ============================================================================

def squash_file(path: str) -> str:
    """Read file and collapse every whitespace run to a single space."""
    if not os.path.isfile(path):
        return ""
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        return re.sub(r"\s+", " ", f.read())


def read_lines(path: str) -> list:
    """Return list of text lines from path, or [] on failure."""
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            return f.readlines()
    except OSError:
        return []


def find_file_under(root: str, filename: str) -> str:
    """
    Case-insensitive search for filename under root.
    Returns the first full path found, or ''.
    """
    target = filename.lower()
    for dirpath, _dirs, files in os.walk(root):
        for fn in files:
            if fn.lower() == target:
                return os.path.join(dirpath, fn)
    return ""


# ============================================================================
# MISSING-HEADER DETECTION
# ============================================================================

_MISSING_RE = re.compile(
    r"fatal\s+error\s*:\s*"
    r"(?:[^\s:]+[\\/])?"
    r"([^\s/\\:]+\.h)"
    r"\s*:\s*"
    r"No\s+such\s+file\s+or\s+directory",
    re.IGNORECASE,
)


def extract_missing_headers(*log_paths) -> list:
    missing = set()
    for path in log_paths:
        for m in _MISSING_RE.finditer(squash_file(path)):
            hdr = m.group(1)
            missing.add(hdr)
            log(f"    [MISSING HEADER] {hdr}", also_print=False)
    return sorted(missing)


# ============================================================================
# MACRO ERROR DETECTION & AUTO-FIX ENGINE
# ============================================================================
#
# BUG FIXES vs previous version:
#  1. Log is squashed before regex so wrapped lines are matched.
#  2. Regex anchored to just the basename (VectorCAST prefixes paths
#     with indentation that becomes spaces after squash).
#  3. _find_all_macro_definitions returns ALL occurrences, not just the first.
#     We look for the one inside an #else block.
#  4. CCAST_.CFG is fully rewritten each attempt from DEFINES[], so we never
#     lose or duplicate entries across retries.

# After squashing, error lines appear in two formats:
#
#   VectorCAST internal log (pipe-separated):
#     "... App_Variable.h : 2818 | 18: error: expected ';' before 'void' ..."
#
#   GCC / vcqik.ERR direct output (colon-separated, optional column):
#     "App_Variable.h:2818:5: error: expected ';' before 'void'"
#     "App_Variable.h:2818: error: expected ';' before 'void'"
#
# The regex below covers both formats.
# Capture: group(1)=filename, group(2)=line number
_MACRO_ERR_RE = re.compile(
    r"([\w.]+\.(?:h|c|cpp))"          # filename (basename only after squash)
    r"\s*[:|]\s*(\d+)"                  # :line  or  | line
    r"(?:\s*[:|]\s*\d+)?"              # optional :col  (gcc format)
    r"\s*[:|]\s*(?:\d+\s*[:|]\s*)?"   # optional VectorCAST pipe col
    r"error\s*:\s*"
    r"expected\s+\S+\s+before\s+\S+",
    re.IGNORECASE,
)

# ── #error "Please specify compiler." detector ───────────────────────────────
# Matches lines like:
#   Crypto_76_HaeModule_MemMap.h:88: error: #error "Please specify compiler."
# Captures group(1)=filename, group(2)=line number
_HASH_ERROR_RE = re.compile(
    r"([\w.]+\.h)"
    r"\s*[:|]\s*(\d+)"
    r"(?:\s*[:|]\s*\d+)?"
    r"\s*[:|]\s*(?:\d+\s*[:|]\s*)?"
    r'error\s*:\s*#\s*error\b[^"]*"([^"]*)"',
    re.IGNORECASE,
)

# Any ALL-CAPS identifier followed by (
_MACRO_CALL_RE = re.compile(r"\b([A-Z][A-Z0-9_]{2,})\s*\(")

# #define MACRONAME(  – captures name
_DEFINE_RE = re.compile(r"^\s*#\s*define\s+([A-Z][A-Z0-9_]*)\s*\(")

# All defined(GUARD) occurrences on one line
_DEFINED_ARGS_RE = re.compile(r"defined\s*\(\s*([A-Z0-9_]+)\s*\)", re.IGNORECASE)

# #ifdef GUARD  (no parentheses)
_IFDEF_RE = re.compile(r"^\s*#\s*ifdef\s+([A-Z0-9_]+)", re.IGNORECASE)


def _classify_define_block(lines: list, def_idx: int):
    """
    Given all lines of a file and the 0-based index of a #define line,
    walk backwards to determine whether it sits in an #if or #else block.

    Returns (block_type, guard_macros) where:
      block_type   : 'if_branch' | 'else_branch' | 'top_level'
      guard_macros : list of macro names from the controlling #if condition
    """
    depth = 0
    in_else = False
    for back in range(def_idx - 1, -1, -1):
        bline = lines[back].strip()

        if bline.startswith("#endif"):
            depth += 1                       # nested block – skip its opener
        elif re.match(r"^#\s*if", bline) and depth > 0:
            depth -= 1                       # matched a nested #if – skip it
        elif bline.startswith("#else") and depth == 0:
            in_else = True                   # we are in the #else arm
        elif re.match(r"^#\s*if", bline) and depth == 0:
            # This is the controlling #if
            guards = _DEFINED_ARGS_RE.findall(bline)
            if not guards:
                m = _IFDEF_RE.match(bline)
                if m:
                    guards = [m.group(1)]
            block_type = "else_branch" if in_else else "if_branch"
            return block_type, guards

    return "top_level", []


def _find_all_macro_definitions(macro_name: str, search_root: str) -> list:
    """
    Return list of (file_path, def_line_idx, block_type, guard_macros)
    for EVERY #define of macro_name found anywhere under search_root.
    """
    results = []
    for dirpath, _dirs, filenames in os.walk(search_root):
        for fn in filenames:
            if not fn.lower().endswith((".h", ".c", ".cpp")):
                continue
            fpath = os.path.join(dirpath, fn)
            lines = read_lines(fpath)
            for idx, line in enumerate(lines):
                m = _DEFINE_RE.match(line)
                if m and m.group(1) == macro_name:
                    block_type, guard_macros = _classify_define_block(lines, idx)
                    results.append((fpath, idx, block_type, guard_macros))
    return results


def resolve_macro_errors(log_paths: tuple, search_root: str,
                         existing_defines: list) -> list:
    """
    Scan logs for parse errors, backtrace macro definitions, and return
    a list of new guard macro names that should be added as -D defines.
    """
    # Squash all logs together
    combined = " ".join(squash_file(p) for p in log_paths)

    errors = []
    seen = set()
    for m in _MACRO_ERR_RE.finditer(combined):
        key = (m.group(1), int(m.group(2)))   # (filename, lineno)
        if key not in seen:
            seen.add(key)
            errors.append(key)

    if not errors:
        log("  [MACRO] No macro-parse errors detected in logs.")
        return []

    log(f"  [MACRO] {len(errors)} parse-error location(s) found.")
    new_defines = []

    for fname, lineno in errors:
        log(f"  [MACRO] Error location: {fname}:{lineno}")

        # Find the file on disk (only need basename match)
        real_path = find_file_under(search_root, fname)
        if not real_path:
            log(f"  [MACRO][WARN] File not found on disk: {fname}")
            continue

        lines = read_lines(real_path)
        if lineno < 1 or lineno > len(lines):
            log(f"  [MACRO][WARN] Line {lineno} out of range ({len(lines)} lines) in {fname}")
            continue

        src_line = lines[lineno - 1]
        log(f"  [MACRO] Offending line {lineno}: {src_line.rstrip()}")

        macros = _MACRO_CALL_RE.findall(src_line)
        log(f"  [MACRO] Macros on that line: {macros}")

        for macro in macros:
            all_defs = _find_all_macro_definitions(macro, search_root)
            if not all_defs:
                log(f"  [MACRO] No definition found for {macro}")
                continue

            log(f"  [MACRO] {macro} has {len(all_defs)} definition(s):")
            for fpath, didx, btype, guards in all_defs:
                log(f"    {fpath}:{didx+1}  block={btype}  guards={guards}")

            # ── KEY LOGIC ────────────────────────────────────────────
            # We want the guard from the #if so we can activate the #if
            # branch and skip the broken #else branch.
            #
            # Case A: macro has an #else definition (the broken one).
            #         Find the corresponding #if definition → take its guards.
            # Case B: macro is only defined in an #if branch but guard is missing.
            #         Take that guard.
            # Case C: top-level only → nothing to do.

            else_defs = [(f, i, bt, g) for f, i, bt, g in all_defs if bt == "else_branch"]
            if_defs   = [(f, i, bt, g) for f, i, bt, g in all_defs if bt == "if_branch"]

            if else_defs:
                # Find a matching #if def in the same file as the #else def
                for ef, ei, _, _ in else_defs:
                    matching_if = [d for d in if_defs if d[0] == ef]
                    if matching_if:
                        guards = matching_if[0][3]
                    else:
                        # Fall back: use the first #if def in any file
                        guards = if_defs[0][3] if if_defs else []
                    for g in guards:
                        if g not in existing_defines and g not in new_defines:
                            log(f"  [MACRO][FIX] Adding -D{g}  "
                                f"(activates #if branch of {macro}, suppresses broken #else)")
                            new_defines.append(g)
                    break   # one else_def is enough to find the guard

            elif if_defs:
                # Macro only defined in #if but guard not active
                for _, _, _, guards in if_defs:
                    for g in guards:
                        if g not in existing_defines and g not in new_defines:
                            log(f"  [MACRO][FIX] Adding -D{g}  "
                                f"(activates #if-only definition of {macro})")
                            new_defines.append(g)
            else:
                log(f"  [MACRO] {macro} is top-level in all files – skipping.")

    return new_defines


# ============================================================================
# #error "Please specify compiler." AUTO-FIX ENGINE
# ============================================================================
#
# Pattern seen in AUTOSAR MCAL / crypto module headers:
#
#   #if defined(COMPILER_GHS)
#     #define CRYPTO_START_SEC_CODE
#     #include "Crypto_76_HaeModule_MemMap.h"
#   #elif defined(COMPILER_IAR)
#     ...
#   #else
#     #error "Please specify compiler."    ← triggered when no guard active
#   #endif
#
# Strategy:
#   1. Find every   filename:line: error: #error "..."  in the logs.
#   2. Open that file on disk, go to that line.
#   3. Walk backwards from the #error line to find the controlling #if.
#   4. Collect ALL defined(GUARD) / #ifdef GUARD identifiers from that #if chain.
#   5. Pick the first guard that looks like a compiler selector
#      (contains "COMPILER", "GCC", "MINGW", "GNU", "GHS", "IAR", "TASKING",
#       "GREEN_HILLS", "WINDRIVER", or "COSMIC") — that is the one MinGW needs.
#   6. If no compiler-shaped guard is found, return all guards and let the
#      caller decide.

_COMPILER_GUARD_KEYWORDS = re.compile(
    r"COMPILER|GCC|MINGW|GNU|GHS|IAR|TASKING|GREEN_HILLS|WINDRIVER|COSMIC|LLVM|ARMCC",
    re.IGNORECASE,
)


def _walk_back_to_if_guards(lines: list, error_line_idx: int) -> list:
    """
    Starting at error_line_idx (0-based), walk backwards through preprocessor
    directives to collect the guards from the outermost controlling #if / #elif
    chain.

    Returns a flat list of guard macro names found in defined() or #ifdef.
    """
    depth   = 0
    guards  = []
    seen_else = False

    for back in range(error_line_idx - 1, -1, -1):
        bline = lines[back].strip()

        if not bline.startswith("#"):
            continue

        if bline.startswith("#endif"):
            depth += 1
            continue

        if re.match(r"^#\s*if", bline) and depth > 0:
            depth -= 1
            continue

        if depth > 0:
            continue

        # depth == 0 from here
        if bline.startswith("#else") and not seen_else:
            seen_else = True
            continue

        if re.match(r"^#\s*elif", bline):
            # Collect guards from every #elif branch too
            guards += _DEFINED_ARGS_RE.findall(bline)
            m_ifdef = _IFDEF_RE.match(bline)
            if m_ifdef:
                guards.append(m_ifdef.group(1))
            continue

        if re.match(r"^#\s*if\b", bline) and depth == 0:
            # Outermost controlling #if — collect its guards and stop
            guards += _DEFINED_ARGS_RE.findall(bline)
            m_ifdef = _IFDEF_RE.match(bline)
            if m_ifdef:
                guards.append(m_ifdef.group(1))
            break

    return guards


def resolve_hash_errors(log_paths: tuple, search_root: str,
                        existing_defines: list) -> list:
    """
    Scan logs for   filename:line: error: #error "..."   messages.
    Backtrace the file to find the compiler-selector guard that should be
    defined so that MinGW falls into the right #if branch instead of hitting
    the #else/#error.

    Returns a list of new guard macro names to add as -D defines.
    """
    combined = " ".join(squash_file(p) for p in log_paths)

    # Collect unique (filename, lineno, message) triples
    hash_errors: list = []
    seen: set = set()
    for m in _HASH_ERROR_RE.finditer(combined):
        key = (m.group(1).lower(), int(m.group(2)))
        if key not in seen:
            seen.add(key)
            hash_errors.append((m.group(1), int(m.group(2)), m.group(3)))

    if not hash_errors:
        log("  [HASH_ERR] No #error directive failures detected.")
        return []

    log(f"  [HASH_ERR] {len(hash_errors)} #error location(s) found.")
    new_defines: list = []

    for fname, lineno, msg in hash_errors:
        log(f"  [HASH_ERR] {fname}:{lineno}  msg='{msg}'")

        real_path = find_file_under(search_root, fname)
        if not real_path:
            log(f"  [HASH_ERR][WARN] File not found on disk: {fname}")
            continue

        lines = read_lines(real_path)
        if lineno < 1 or lineno > len(lines):
            log(f"  [HASH_ERR][WARN] Line {lineno} out of range in {fname}")
            continue

        log(f"  [HASH_ERR] Offending line: {lines[lineno - 1].rstrip()}")

        guards = _walk_back_to_if_guards(lines, lineno - 1)   # 0-based
        log(f"  [HASH_ERR] Candidate guards from #if chain: {guards}")

        if not guards:
            log(f"  [HASH_ERR][WARN] Could not find controlling #if for {fname}:{lineno}")
            continue

        # Prefer a guard that looks like a compiler selector
        compiler_guards = [g for g in guards if _COMPILER_GUARD_KEYWORDS.search(g)]
        chosen = compiler_guards if compiler_guards else guards

        for g in chosen:
            if g not in existing_defines and g not in new_defines:
                log(f"  [HASH_ERR][FIX] Adding -D{g}  "
                    f"(satisfies compiler guard in {fname}:{lineno})")
                new_defines.append(g)

    return new_defines

def find_header_dirs(root: str, headers: list) -> dict:
    needed  = set(headers)
    found   = {}
    lowered = {h.lower(): h for h in needed}

    log(f"  [SEARCH] Walking: {root}")
    log(f"  [SEARCH] Looking for: {sorted(needed)}")

    for dirpath, _dirs, filenames in os.walk(root):
        files_lower = {fn.lower() for fn in filenames}
        for lh, orig in list(lowered.items()):
            if lh in files_lower and orig not in found:
                found[orig] = dirpath
                log(f"  [FOUND]  {orig}  ->  {dirpath}")
        if set(found.keys()) == needed:
            break

    for h in sorted(needed - set(found.keys())):
        log(f"  [NOT FOUND]  {h}")

    return found


# ============================================================================
# CCAST_.CFG WRITER  –  always a full rewrite so DEFINES[] is the single source
# ============================================================================

def write_cfg() -> None:
    """
    Write CCAST_.CFG from scratch using the current DEFINES list.
    Called before every build attempt so newly discovered defines are included.
    """
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
    ]
    with open(CFG_FILE, "w", encoding="utf-8") as f:
        f.write("\n".join(cfg_lines) + "\n")
        for d in DEFINES:
            f.write(f"C_DEFINE_LIST: {d}\n")
    log(f"  [CFG] Defines active ({len(DEFINES)}): {' '.join(DEFINES)}")


# ============================================================================
# ENV SCRIPT WRITER
# ============================================================================

def write_env_script(extra_includes: list) -> None:
    lines = [
        "-- VectorCAST Environment Script (auto-generated)",
        f"-- Environment: {ENV_NAME}",
        "--",
        "ENVIRO.NEW",
        f"ENVIRO.NAME: {ENV_NAME}",
        "ENVIRO.COVERAGE_TYPE: Statement+MCDC",
        "ENVIRO.WHITE_BOX: YES",
        "ENVIRO.STUB: ALL_BY_PROTOTYPE",
        "ENVIRO.COMPILER: CC",
        "ENVIRO.TYPE_HANDLED_DIRS_ALLOWED:",
        f"ENVIRO.UUT: {UUT_FILE}",
        f"ENVIRO.BASE_DIRECTORY: {BASE_DIR_NAME}={BASE_DIR_PATH}",
    ]
    for src in [SOURCE_DIR_1, SOURCE_DIR_2, SOURCE_DIR_3]:
        if src:
            norm_base = os.path.normcase(os.path.normpath(BASE_DIR_PATH))
            norm_src  = os.path.normcase(os.path.normpath(src))
            if norm_src.startswith(norm_base):
                rel = os.path.normpath(src)[len(os.path.normpath(BASE_DIR_PATH)):].lstrip(os.sep)
                entry = f"$({BASE_DIR_NAME})\\{rel}"
            else:
                entry = src
            lines.append(f"ENVIRO.SEARCH_LIST: {entry}")
    for inc in extra_includes:
        # Convert absolute path to $(BASE_DIR_NAME)\relative form when possible
        # so VectorCAST resolves it the same way as the manually built .env
        norm_base = os.path.normcase(os.path.normpath(BASE_DIR_PATH))
        norm_inc  = os.path.normcase(os.path.normpath(inc))
        if norm_inc.startswith(norm_base):
            rel = os.path.normpath(inc)[len(os.path.normpath(BASE_DIR_PATH)):].lstrip(os.sep)
            entry = f"$({BASE_DIR_NAME})\\{rel}"
        else:
            entry = inc
        lines.append(f"ENVIRO.SEARCH_LIST: {entry}")

    with open(ENV_SCRIPT, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")

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
    with open(ENV_SCRIPT, "a", encoding="utf-8") as f:
        f.write(user_globals)


# ============================================================================
# BUILD RUNNER
# ============================================================================

def collect_vcqik_errors(env_dir: str) -> None:
    """
    After a failed build VectorCAST writes the real compiler diagnostics into
    vcqik/*.vcqik.ERR files inside the environment directory.  The main build
    log only says "See error log for complete message", so the header/macro
    scanners never see the actual error lines unless we collect them here.

    This function reads every .vcqik.ERR file and appends its content to
    ERROR_LOG so that extract_missing_headers(), resolve_macro_errors(), and
    resolve_hash_errors() can all find them in the normal post-build scan.
    """
    vcqik_dir = os.path.join(env_dir, "vcqik")
    if not os.path.isdir(vcqik_dir):
        return

    appended = 0
    with open(ERROR_LOG, "a", encoding="utf-8", errors="replace") as out:
        for fname in sorted(os.listdir(vcqik_dir)):
            if not fname.lower().endswith(".err"):
                continue
            fpath = os.path.join(vcqik_dir, fname)
            lines = read_lines(fpath)
            if not lines:
                continue
            out.write(f"\n\n=== vcqik/{fname} ===\n")
            out.writelines(lines)
            appended += 1

    if appended:
        log(f"  [VCQIK] Appended {appended} .vcqik.ERR file(s) to error_log.txt for scanning")


def run_build(env_dir: str) -> int:
    # FIX: collect vcqik ERR files from the PREVIOUS attempt BEFORE wiping
    # the env directory.  The retry loop calls run_build at the top of every
    # attempt, so without this the detailed compiler errors are lost before
    # extract_missing_headers / resolve_macro_errors can read them.
    if os.path.isdir(env_dir):
        collect_vcqik_errors(env_dir)
        log("  Cleaning previous environment directory...")
        shutil.rmtree(env_dir)

    cmd = [CLICAST_EXE, "-lc", "ENvironment", "Build", ENV_SCRIPT]
    log(f"  CMD: {' '.join(cmd)}")

    with open(BUILD_LOG, "w", encoding="utf-8") as bout, \
         open(ERROR_LOG, "w", encoding="utf-8") as berr:
        result = subprocess.run(cmd, stdout=bout, stderr=berr)

    # FIX: also collect vcqik ERR files from THIS attempt immediately after
    # the build so the very first failure is also diagnosed correctly.
    if result.returncode != 0 and os.path.isdir(env_dir):
        collect_vcqik_errors(env_dir)

    return result.returncode


# ============================================================================
# ENABLE FUNCTION COVERAGE (post-build)
# ============================================================================

def enable_function_coverage() -> None:
    log("\n[FUNCTION COVERAGE] Enabling Function + FunctionCall instrumentation...")

    for key, label in [
        ("VCAST_INSTRUMENT_FOR_FUNCTION_COVERAGE",      "Function"),
        ("VCAST_INSTRUMENT_FOR_FUNCTION_CALL_COVERAGE", "FunctionCall"),
    ]:
        cmd = [CLICAST_EXE, "-lc", "-e", ENV_NAME, "options", "Coverage", key, "TRUE"]
        log(f"  CMD: {' '.join(cmd)}")
        r = subprocess.run(cmd, capture_output=True, text=True)
        if r.returncode != 0:
            log(f"  [WARN] {label} coverage option exit {r.returncode}: {r.stderr.strip()}")

    ri_log = os.path.join(WORK_DIR, "reinstrument_log.txt")
    cmd_ri = [CLICAST_EXE, "-lc", "-e", ENV_NAME, "ENvironment", "Re_instrument"]
    log(f"  CMD: {' '.join(cmd_ri)}")
    with open(ri_log, "w", encoding="utf-8") as rout:
        r3 = subprocess.run(cmd_ri, stdout=rout, stderr=rout)

    ri_text = "".join(read_lines(ri_log)).strip()
    if ri_text:
        print(ri_text)

    if r3.returncode != 0:
        log(f"  [WARN] Re-instrumentation exit {r3.returncode}. Check: {ri_log}")
    else:
        log("  [OK] Function + FunctionCall coverage active.")
        print("  [OK] Function + FunctionCall coverage enabled.")


# ============================================================================
# MAIN
# ============================================================================

def main() -> None:
    sep = "=" * 76

    print(sep)
    print("  VectorCAST Auto-Compile  |  Header + Macro Auto-Resolution")
    print(f"  Environment   : {ENV_NAME}")
    print(f"  VectorCAST    : {VECTORCAST_DIR}")
    print(f"  Working Dir   : {WORK_DIR}")
    print(f"  Search Root   : {HEADER_SEARCH_ROOT}")
    print(f"  Max retries   : {MAX_RETRY_ROUNDS}")
    print(sep)

    # Init detailed log
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    os.makedirs(WORK_DIR, exist_ok=True)
    with open(DETAILED_LOG, "w", encoding="utf-8") as f:
        f.write(f"{sep}\nVectorCAST Auto-Compile Log\nStarted: {now}\n{sep}\n\n")
        f.write(f"  VECTORCAST_DIR : {VECTORCAST_DIR}\n")
        f.write(f"  ENV_NAME       : {ENV_NAME}\n")
        f.write(f"  WORK_DIR       : {WORK_DIR}\n")
        f.write(f"  BASE_DIR_PATH  : {BASE_DIR_PATH}\n")
        f.write(f"  UUT_FILE       : {UUT_FILE}\n")
        f.write(f"  DEFINES        : {' '.join(DEFINES)}\n\n")

    # STEP 0 – verify VectorCAST
    log("\n[STEP 0] Checking VectorCAST...")
    if not os.path.isfile(CLICAST_EXE):
        log(f"[ERROR] clicast.exe not found: {CLICAST_EXE}")
        input("Press Enter to exit...")
        sys.exit(1)
    log(f"  [OK] {CLICAST_EXE}")

    mingw_bin = os.path.join(VECTORCAST_DIR, "MinGW", "bin")
    os.environ["PATH"] = f"{VECTORCAST_DIR};{mingw_bin};{os.environ.get('PATH', '')}"
    os.environ.pop("VECTORCAST_DIR", None)

    os.chdir(WORK_DIR)
    log(f"  [OK] Working dir: {os.getcwd()}")

    # STEP 1 – write initial CCAST_.CFG
    log("\n[STEP 1] Writing CCAST_.CFG...")
    write_cfg()

    # Seed static extra-includes
    extra_includes: list = [
        ei for ei in [EXTRA_INCLUDE_1, EXTRA_INCLUDE_2, EXTRA_INCLUDE_3] if ei
    ]
    already_searched: set  = set()
    already_defines_tried: set = set()   # guards we've already attempted
    env_dir = os.path.join(WORK_DIR, ENV_NAME)

    # ================================================================
    #  AUTO-RETRY LOOP
    # ================================================================
    for attempt in range(1, MAX_RETRY_ROUNDS + 1):
        print()
        print(f"{'─'*76}")
        print(f"  BUILD ATTEMPT {attempt} / {MAX_RETRY_ROUNDS}")
        if DEFINES:
            print(f"  Active defines ({len(DEFINES)}): {' '.join(DEFINES)}")
        if extra_includes:
            print(f"  Include paths  ({len(extra_includes)}):")
            for p in extra_includes:
                print(f"    {p}")
        print(f"{'─'*76}\n")
        log(f"\n{'─'*76}\n[ATTEMPT {attempt}]")

        # Always rewrite CCAST_.CFG from DEFINES[] so nothing is lost or doubled
        write_cfg()

        write_env_script(extra_includes)
        log(f"  [OK] {ENV_NAME}.env written")

        build_exit = run_build(env_dir)
        log(f"  Exit code: {build_exit}")

        # Print build output
        print("=== BUILD OUTPUT ===")
        print("".join(read_lines(BUILD_LOG)))
        err_lines = read_lines(ERROR_LOG)
        if err_lines:
            print("=== ERROR OUTPUT ===")
            print("".join(err_lines))

        # ── SUCCESS ──────────────────────────────────────────────────
        if build_exit == 0:
            log(f"\n[SUCCESS] Built on attempt {attempt}!")
            print(f"\n[SUCCESS] Environment built successfully! (attempt {attempt})")
            if len(DEFINES) > 1:
                print(f"  Auto-added defines: {[d for d in DEFINES if d != '__USE_MINGW_ANSI_STDIO']}")
            if extra_includes:
                print(f"  Auto-added include paths:")
                for p in extra_includes:
                    print(f"    {p}")
            print(f"\n  Environment : {env_dir}")
            print(f"  Build log   : {BUILD_LOG}")
            enable_function_coverage()
            show_alert(
                "VectorCAST Build Success",
                f"SUCCESSFUL after {attempt} attempt(s)!\\n\\n"
                f"Environment: {ENV_NAME}\\nLocation: {env_dir}\\n\\n"
                f"Coverage: Statement+MC/DC+Function+FunctionCall",
            )
            break

        # ── FAILURE ──────────────────────────────────────────────────
        log(f"\n[ATTEMPT {attempt}] Failed.")

        # 1. Try missing-header fix first
        log("  Scanning for missing headers...")
        all_missing  = extract_missing_headers(BUILD_LOG, ERROR_LOG)
        new_headers  = [h for h in all_missing if h not in already_searched]
        log(f"  Missing headers (new): {new_headers}")

        if new_headers:
            already_searched.update(new_headers)
            found_map = find_header_dirs(HEADER_SEARCH_ROOT, new_headers)
            if not found_map:
                fail(
                    f"Could not find {new_headers} under '{HEADER_SEARCH_ROOT}'.\n"
                    "Try setting HEADER_SEARCH_ROOT to a higher-level directory."
                )
            norm_existing = [os.path.normcase(os.path.normpath(p)) for p in extra_includes]
            added = []
            for hdr, dirpath in found_map.items():
                norm = os.path.normcase(os.path.normpath(dirpath))
                if norm not in norm_existing:
                    extra_includes.append(dirpath)
                    norm_existing.append(norm)
                    added.append((hdr, dirpath))
                    log(f"  [+INCLUDE] {dirpath}  (provides {hdr})")
            if added:
                print(f"\n[HEADER-FIX] {len(added)} new include path(s) added:")
                for hdr, p in added:
                    print(f"    {hdr}  ->  {p}")
            print("  Retrying build...\n")
            continue

        # 2. No missing headers – try macro-error fix
        log("  No missing-header errors. Trying macro-error resolver...")
        new_defines = resolve_macro_errors(
            (BUILD_LOG, ERROR_LOG),
            HEADER_SEARCH_ROOT,
            DEFINES,
        )
        # Filter out any guard we've already tried (avoid infinite loop)
        new_defines = [d for d in new_defines if d not in already_defines_tried]

        if new_defines:
            already_defines_tried.update(new_defines)
            print(f"\n[MACRO-FIX] {len(new_defines)} platform guard(s) detected:")
            for d in new_defines:
                print(f"  * -D{d}")
                DEFINES.append(d)
            log(f"  [MACRO-FIX] DEFINES now: {DEFINES}")
            print("  Retrying build with updated defines...\n")
            continue

        # 3. No macro errors – try #error "Please specify compiler." fix
        log("  No macro-parse errors. Trying #error compiler-guard resolver...")
        new_hash_defines = resolve_hash_errors(
            (BUILD_LOG, ERROR_LOG),
            HEADER_SEARCH_ROOT,
            DEFINES,
        )
        new_hash_defines = [d for d in new_hash_defines if d not in already_defines_tried]

        if new_hash_defines:
            already_defines_tried.update(new_hash_defines)
            print(f"\n[HASH-ERROR-FIX] {len(new_hash_defines)} compiler guard(s) detected:")
            for d in new_hash_defines:
                print(f"  * -D{d}")
                DEFINES.append(d)
            log(f"  [HASH-ERROR-FIX] DEFINES now: {DEFINES}")
            print("  Retrying build with updated defines...\n")
            continue

        # 4. No fix worked – extract compiler errors and give up
        log("[ERROR] No header fix, macro fix, or #error compiler-guard fix resolved the failure.")
        if os.path.isdir(env_dir):
            ce_file = os.path.join(WORK_DIR, "compile_errors.txt")
            subprocess.run(
                [CLICAST_EXE, "-e", ENV_NAME,
                 "ENvironment", "Extract", "Compile_errors", ce_file],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            )
            if os.path.isfile(ce_file):
                print("\n--- COMPILER ERRORS ---")
                print("".join(read_lines(ce_file)))
        fail("Build failed. Check detailed_log.txt for the full backtrace.")

    else:
        fail(f"Did not succeed after {MAX_RETRY_ROUNDS} attempts.")

    # ================================================================
    #  SUMMARY
    # ================================================================
    print()
    print(sep)
    print("  BUILD SUMMARY")
    print(sep)
    print(f"  Compiler    : VectorCAST MinGW (C)")
    print(f"  Environment : {ENV_NAME}")
    print(f"  Coverage    : Statement+MC/DC+Function+FunctionCall  |  Whitebox: YES")
    print(f"  UUT         : {UUT_FILE}.c")
    print(f"  Working Dir : {WORK_DIR}")
    print(f"  Defines     : {' '.join(DEFINES)}")
    if extra_includes:
        print(f"  Auto-includes ({len(extra_includes)}):")
        for p in extra_includes:
            print(f"    {p}")
    print(sep)
    log(f"\n{sep}\nSUMMARY – {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    log(f"Defines ({len(DEFINES)}): {DEFINES}")
    log(f"Includes ({len(extra_includes)}): {extra_includes}")
    print("\nDone.")


if __name__ == "__main__":
    main()
