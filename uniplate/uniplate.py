#!/usr/bin/env python3
from __future__ import print_function
import argparse
import os.path
import importlib


parser = argparse.ArgumentParser()
parser.add_argument("template", help="Template to use (should work with any open document, tested on .odg)", type=str)
parser.add_argument("table", help="Table with data (.ods file)", type=str)
parser.add_argument("-o", "--outdir", help="Output directory", default=".")
parser.add_argument("-v", "--verbosity", action="count", default=0, help="Be verbose")
parser.add_argument("-n", "--naming", default=None, type=str, help="File naming pattern, like {field1}_{field2}")
parser.add_argument("-s", "--sheet", default=[], type=str, action='append',
                    help="Sheet to process (can occur more than once)")
parser.add_argument("-e", "--skip-empty", default=False, action='store_true',
                    help="Skip empty values in key:value columns")
parser.add_argument("-f", "--fill-with-last", default=False, action='store_true',
                    help="Fill key:value columns with last pair instead of blank ")
parser.add_argument("--table-loader", default="uniplate.uniplate_engine", type=str,
                    help="Table loader module (default is uniplate_engine")
parser.add_argument("--templater", default="uniplate.uniplate_engine", type=str,
                    help="Templater module (default is uniplate_engine")
parser.parse_args()

args = parser.parse_args()

# Fool protection

if not os.path.isdir(args.outdir):
    print("{} is not a directory".format(args.outdir))
    exit()

if not os.path.isfile(args.template):
    print("{} is not a file".format(args.template))
    exit()

if not os.path.isfile(args.table):
    print("{} is not a file".format(args.table))
    exit()

# Parse the table

TableLoader = importlib.import_module(args.table_loader).TableLoader
Templater = importlib.import_module(args.templater).Templater

table, naming = TableLoader.load_table(args)

if args.verbosity > 1:
    print("Table opened")

# Load the template

template_object = Templater(args, naming)

if args.verbosity > 1:
    print("Template loaded")

if args.verbosity > 1:
    print("Table loaded")

# Template and save each row

if args.verbosity > 1:
    print("Processing entries")

for row in table:
    template_object.template_file(row)
