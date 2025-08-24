#!/usr/bin/env python
# -*- coding: utf-8 -*-
import logging
import argparse
import json
import os


from src.paperwork import create_printable_sheets
import src.paperwork

'''
    # Reading an Excel file
    print(args.input_file)
    wb = load_workbook(args.input_file)
    ws = wb.active

    # Reading data from a specific cell
    data = ws['A1'].value
    print("A1 = ", data)
'''

from xlcalculator import ModelCompiler
from xlcalculator import Model
from xlcalculator import Evaluator



def main(args):
    logging.basicConfig(format='%(levelname)-5s %(asctime)s: %(message)s', datefmt='%H:%M:%S', level=logging.DEBUG)


    compiler = ModelCompiler()
    new_model = compiler.read_and_parse_archive(args.input_file)
    src.paperwork.evaluator = Evaluator(new_model)


    create_printable_sheets(args.input_file, "output.xlsx")



if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input_file', help='Output file name')
    args = parser.parse_args()

    main(args)
