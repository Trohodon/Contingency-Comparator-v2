import clr
import sys
from pathlib import Path
import System

DLL_PATH = Path(r"\\mbu.ad.dominionnet.com\data\TRANSMISSION OPERATIONS CENTER\7T\Data2\DESC_Trans_Planning\LTR_General\SOFTWARE\_IN HOUSE\TLICs\bin\tliclib.dll")
# ^ make sure this points exactly to the dll file

# Make sure the DLL folder is on sys.path (helps dependency resolution)
sys.path.append(str(DLL_PATH.parent))

# Add reference (good practice, though reflection will still work without it)
clr.AddReference(str(DLL_PATH))

# Load assembly for reflection
asm = System.Reflection.Assembly.LoadFile(str(DLL_PATH))

# Find the type safely
tline_type = None
for t in asm.GetTypes():
    if t.FullName == "TLICLib.TLine":
        tline_type = t
        break

print("tline_type =", tline_type)

if tline_type is None:
    raise Exception("Could not find TLICLib.TLine inside assembly")

print("\nMethods on TLine:\n")
for m in tline_type.GetMethods():
    print(m.Name)