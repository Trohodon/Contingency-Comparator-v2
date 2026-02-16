"""
TLIC DLL Inspector (pythonnet)

What this does:
- Verifies the DLL path is real (exists, size, timestamp, hash)
- Loads the DLL in a reliable way (LoadFrom)
- Prints assembly identity + load location
- Lists all types inside the DLL
- Locates TLICLib.TLine
- Prints constructors, properties, and methods for TLine (and optionally other classes)

Install:
    pip install pythonnet

Run:
    python tlic_dll_inspector.py
"""

import sys
import hashlib
from pathlib import Path

import clr  # pythonnet
import System


# =========================
# CONFIG: SET YOUR DLL PATH
# =========================
DLL_PATH = Path(
    r"\\mbu.ad.dominionnet.com\data\TRANSMISSION OPERATIONS CENTER\7T\Data2\DESC_Trans_Planning\LTR_General\SOFTWARE\_IN HOUSE\TLICs\bin\tliclib.dll"
)

# If you want to keep a local copy for stability, set this instead:
# DLL_PATH = Path(r"C:\Users\isaak01\source\repos\TLIC_Remake\libs\tliclib.dll")


# =========================
# Helpers
# =========================
def sha256_first_mb(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        h.update(f.read(1024 * 1024))
    return h.hexdigest()


def dump_loaded_tliclib_assemblies():
    loaded = []
    for a in System.AppDomain.CurrentDomain.GetAssemblies():
        try:
            name = a.GetName().Name
            if name and name.lower() == "tliclib":
                loaded.append((a.FullName, a.Location))
        except Exception:
            # Some dynamic assemblies can throw on Location
            loaded.append((str(a.FullName), "<no location>"))
    print("\nAlready-loaded 'tliclib' assemblies in this process:")
    if not loaded:
        print("  (none)")
    else:
        for full, loc in loaded:
            print(f"  {full}\n    {loc}")


def load_assembly(path: Path) -> System.Reflection.Assembly:
    """
    LoadFrom is generally more reliable than LoadFile for type resolution
    because of .NET load contexts.
    """
    # Helps dependency resolution if there are other DLLs next to it
    sys.path.append(str(path.parent))

    # AddReference is fine but not required for reflection;
    # we still do it because you’ll likely want to CALL into types later.
    clr.AddReference(str(path))

    # Load for reflection
    asm = System.Reflection.Assembly.LoadFrom(str(path))
    return asm


def list_types(asm: System.Reflection.Assembly):
    types = asm.GetTypes()
    print("\n--- Types inside DLL ---")
    for t in types:
        print(f"  {t.FullName}")
    return types


def get_type_exact(types, full_name: str):
    for t in types:
        if t.FullName == full_name:
            return t
    return None


def print_type_details(t):
    print(f"\n===== {t.FullName} =====")

    # Constructors
    print("\n-- Constructors --")
    ctors = t.GetConstructors()
    if ctors is None or len(ctors) == 0:
        print("  (none)")
    else:
        for c in ctors:
            print(f"  {c}")

    # Properties
    print("\n-- Properties --")
    props = t.GetProperties()
    if props is None or len(props) == 0:
        print("  (none)")
    else:
        for p in props:
            try:
                pt = p.PropertyType.FullName
            except Exception:
                pt = str(p.PropertyType)
            print(f"  {p.Name}: {pt}")

    # Methods (filter out the huge Object base methods if you want)
    print("\n-- Methods --")
    methods = t.GetMethods()
    if methods is None or len(methods) == 0:
        print("  (none)")
    else:
        # Show unique method names in order (cleaner)
        seen = set()
        for m in methods:
            name = m.Name
            if name in seen:
                continue
            seen.add(name)
            print(f"  {name}")


# =========================
# Main
# =========================
def main():
    print("DLL_PATH =", DLL_PATH)

    # Basic file verification
    print("Exists?  =", DLL_PATH.exists())
    print("Is file? =", DLL_PATH.is_file())
    if not DLL_PATH.exists() or not DLL_PATH.is_file():
        raise FileNotFoundError(f"DLL not found or not a file: {DLL_PATH}")

    st = DLL_PATH.stat()
    print("Size    =", st.st_size)
    print("MTime   =", st.st_mtime)
    print("SHA256(first1MB) =", sha256_first_mb(DLL_PATH))

    # Important: if VS/debugger has already loaded some tliclib, show it
    dump_loaded_tliclib_assemblies()

    # Load assembly
    asm = load_assembly(DLL_PATH)
    print("\nAssembly FullName :", asm.FullName)
    print("Assembly Location :", asm.Location)

    # List types
    types = list_types(asm)

    # Find TLine
    tline = get_type_exact(types, "TLICLib.TLine")
    if tline is None:
        print("\n!!! Could not find TLICLib.TLine in this assembly.")
        print("Tip: This usually means you're pointing at a different tliclib.dll than expected.")
        print("Tip: Look above at the Types list and confirm which namespace/class names exist.")
        return

    # Print TLine details
    print_type_details(tline)

    # Optional: also inspect other key types quickly
    for name in [
        "TLICLib.Branch",
        "TLICLib.Structure",
        "TLICLib.Conductor",
        "TLICLib.Position",
    ]:
        t = get_type_exact(types, name)
        if t is not None:
            print_type_details(t)

    print("\nDone.")


if __name__ == "__main__":
    main()