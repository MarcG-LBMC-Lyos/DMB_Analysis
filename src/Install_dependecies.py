import os

modules = ('matplotlib==3.1.2', 'scikit-image==0.16.2', 'numpy==1.18.1', 'Pillow', 'xlsxwriter', 'xlrd')

os.system('python -m pip install --upgrade pip')
os.system('python -m pip install --upgrade setuptools')
for module in modules:
    os.system('python -m pip install %s'%module)
