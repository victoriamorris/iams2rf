python setup.py install
pyinstaller bin/snapshot2sql.py -F
pyinstaller bin/sql2rf.py -F
read -p "Press [Enter]"
rm -rf bin/__pycache__
mv dist/snapshot2sql.exe snapshot2sql.exe
mv dist/sql2rf.exe sql2rf.exe
rmdir dist
rm -rf __pycache__
rm -rf build
rm snapshot2sql.spec
rm sql2rf.spec
read -p "Press [Enter]"