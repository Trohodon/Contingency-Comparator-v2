py -m PyInstaller --onefile --windowed ^
  --icon=ContingencyComparaterV2\assets\app.ico ^
  --add-data "ContingencyComparaterV2\assets\app.ico;assets" ^
  --add-data "ContingencyComparaterV2\assets\splash.png;assets" ^
  ContingencyComparaterV2\ContingencyComparaterV2.py