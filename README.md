# Contingency-Comparator-v2

cd C:\Users\isaak01\source\repos\ContingencyComparaterV2

py -X importtime ContingencyComparaterV2\ContingencyComparaterV2.py 2> importtime.txt

cd C:\Users\isaak01\source\repos\ContingencyComparaterV2
py -X importtime .\ContingencyComparaterV2\ContingencyComparaterV2.py 2>&1 | out-file -encoding utf8 importtime.txt

cd dist\ContingencyComparaterV2
gci _internal -Directory | % {
  $size = (gci $_.FullName -Recurse -File | measure Length -Sum).Sum
  [pscustomobject]@{Name=$_.Name; MB=[math]::Round($size/1MB,1)}
} | sort-object MB -desc | select-object -first 25