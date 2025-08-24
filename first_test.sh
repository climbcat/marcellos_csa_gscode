#!/bin/sh
#!/bin/sh
virtualenv csaenv
source csaenv/bin/activate
pip install openpyxl
python3 test.py
deactivate
rm -r csaenv
