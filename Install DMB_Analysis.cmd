set installPath=%cd%\python37

.\py\python-3.7.9-amd64.exe /passive TargetDir="%installPath%"

cd .\src
..\python37\python.exe -m pip install "matplotlib==3.1.2"
..\python37\python.exe -m pip install "xlrd==2.0.1"
..\python37\python.exe -m pip install "Pillow"
..\python37\python.exe -m pip install "xlsxwriter==3.1.6"
..\python37\python.exe -m pip install "scikit-image==0.16.2"
..\python37\python.exe -m pip install "numpy==1.18.1"

cd..
echo .\python37\python.exe ".\src\DMB_Analysis.py" > .\DMB_Analysis.cmd