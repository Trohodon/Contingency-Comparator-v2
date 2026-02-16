from pathlib import Path
import hashlib

DLL_PATH = Path(r"\\...your...\tliclib.dll")

print("DLL_PATH =", DLL_PATH)
print("Exists?  =", DLL_PATH.exists())
print("Is file? =", DLL_PATH.is_file())
if DLL_PATH.exists():
    print("Size    =", DLL_PATH.stat().st_size)
    print("MTime   =", DLL_PATH.stat().st_mtime)

    # quick hash (proves it's the same binary between runs)
    h = hashlib.sha256()
    with open(DLL_PATH, "rb") as f:
        h.update(f.read(1024 * 1024))  # first 1MB is enough as a fingerprint
    print("SHA256(first1MB) =", h.hexdigest())