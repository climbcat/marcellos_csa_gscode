#!/bin/sh
virtualenv csaenv
source csaenv/bin/activate
pip install openpyxl
deactivate
