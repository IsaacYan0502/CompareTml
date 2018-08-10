@echo off  
  
echo .
timeout /t 1 > nul
echo .. 
timeout /t 1 > nul
echo ...
timeout /t 1 > nul
echo ....
timeout /t 1 > nul

echo Wait for preparing the environment for python...
timeout /t 1 > nul

pip install xlwt > nul
pip install xlrd > nul
pip install requests > nul
pip install lxml > nul
pip install xlutils > nul
pip install bs4 > nul
pip install progressbar2 > nul

echo Start to run the compare function........
timeout /t 1 > nul

python downloadTmlCompareWithExcelBase.py

pause  

