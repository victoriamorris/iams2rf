python setup.py install
pyinstaller bin/researcherFormat.py -F
pyinstaller bin/write_rf_config.py -F
read -p "Press [Enter]"
rm -rf bin/__pycache__
mv dist/researcherFormat.exe researcherFormat.exe
mv dist/write_rf_config.exe write_rf_config.exe
rmdir dist
rm -rf __pycache__
rm -rf build
rm researcherFormat.spec
rm write_rf_config.spec
read -p "Press [Enter]"