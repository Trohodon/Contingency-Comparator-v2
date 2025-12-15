# Contingency-Comparator-v2

py -m pip install --upgrade pywin32 pywin32-ctypes `
  --trusted-host pypi.org `
  --trusted-host files.pythonhosted.org `
  --trusted-host pypi.python.org

py -m pywin32_postinstall -install

py -m PyInstaller --onefile --windowed ContingencyComparaterV2.py