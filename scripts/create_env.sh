#!/bin/sh
cd "$(dirname "$0")"
virtualenv csaenv
source csaenv/bin/activate
pip install openpyxl
pip install xlcalculator
deactivate
