#!/usr/bin/env python3
from __future__ import print_function
import argparse
import os.path
import zipfile
import odf.opendocument
import odf.table
import odf.text


parser = argparse.ArgumentParser()
parser.add_argument("template", help="Template to use (should work with any open document, tested on .odg)", type=str)
parser.add_argument("table", help="Table with data (.ods file)", type=str)
parser.add_argument("-o", "--outdir", help="Output directory", default=".")
parser.add_argument("-v", "--verbosity", action="count", default=0, help="Be verbose")
parser.add_argument("-n", "--naming", default=None, type=str, help="File naming pattern, like {field1}_{field2}")
parser.add_argument("-s", "--sheet", default=[], type=str, action='append', help="Sheet to process (can occur more than once)")
parser.parse_args()

args = parser.parse_args()


def node_value(node):
    """Tries to convert an Element into text value"""
    text = ''
    if hasattr(node, "childNodes"):
        for n in node.childNodes:
            text += node_value(n)
    if hasattr(node, 'data'):
        text += node.data
    return text


def cell_value(cell):
    """Returns a cell as text value"""
    ps = cell.getElementsByType(odf.text.P)
    text_content = ""
    for p in ps:
        text_content += node_value(p)
    return text_content


def template_string(string, name, text):
    """Substitutes {name} to text in the string"""
    if args.verbosity > 4:
        print(string)
    return string.replace("{"+name+"}", text)


def template_node(node, name, text):
    """Substitutes {name} to text within an Element and children.
     {name} has to be within a single Element."""
    if hasattr(node, "childNodes"):
        for n in node.childNodes:
            template_node(n, name, text)
    if hasattr(node, 'data'):
        if "{" + name + "}" in node.data:
            node.data = template_string(node.data, name, text)


def template(node, name, text):
    """Substitutes {name} to text within a Document"""
    ps = node.getElementsByType(odf.text.P)
    for p in ps:
        template_node(p, name, text)


def load_template():
    """Tries to load a template, exits if the template is corrupted or not existing"""
    try:
        template_object = odf.opendocument.load(args.template)
    except zipfile.BadZipFile:
        print(
            "{} is not a valid zip archive (which means it's also not a .ods document for sure)".format(args.template))
        exit()

    if template_object.mimetype != 'application/vnd.oasis.opendocument.graphics':
        print("{} is not an Open Document Chart".format(args.table))
        exit()

    return template_object


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

template_object = load_template()

if args.verbosity > 5:
    print(template_object.contentxml())

if args.verbosity > 1:
    print("Template loaded")

try:
    table_object = odf.opendocument.load(args.table)
except zipfile.BadZipFile:
    print("{} is not a valid zip archive (which means it's also not a .ods document for sure)".format(args.table))
    exit()

if table_object.mimetype != 'application/vnd.oasis.opendocument.spreadsheet':
    print("{} is not an Open Document Chart".format(args.table))
    exit()

if args.verbosity > 1:
    print("Table opened")


# Parse the table

naming = args.naming

table = []
for sheet in table_object.spreadsheet.getElementsByType(odf.table.Table):
    s_name = sheet.getAttribute("name")
    if args.verbosity > 1:
        print("Processing sheet `{}'".format(s_name))
    rows = sheet.getElementsByType(odf.table.TableRow)

    if len(args.sheet) != 0 and s_name not in args.sheet:
        if args.verbosity > 1:
            print("Skipping sheet `{}' - not requested.".format(s_name))
        continue

    if len(rows) < 2:
        if args.verbosity > 1:
            print("Sheet `{}' is empty.".format(s_name))
            continue
    header_row = rows[0]
    rows = rows[1:]

    header = [cell_value(cell) for cell in header_row.getElementsByType(odf.table.TableCell)]
    if naming is None:
        naming = "{"+header[0]+"}"
    for row in rows:
        row_dictionary = {}
        cells = [cell_value(cell) for cell in row.getElementsByType(odf.table.TableCell)]
        for i in range(len(header)):
            try:
                value = cells[i]
            except IndexError:
                value = ""
            row_dictionary[header[i]] = value
        if args.verbosity > 2:
            print(row_dictionary)
        table.append(row_dictionary)

if args.verbosity > 1:
    print("Table loaded")


# Template and save each row

if args.verbosity > 1:
    print("Processing entries")

for row in table:
    template_object = load_template()
    filename = naming
    for key in row:
        template(template_object, key, row[key])
        filename = template_string(filename, key, row[key])
    if os.path.isfile(os.path.join(args.outdir, filename+".odg")):
        suffix = 1
        while os.path.isfile(os.path.join(args.outdir, filename+"_{}".format(suffix)+".odg")):
            suffix += 1
        filename += "_{}".format(suffix)
    path = os.path.join(args.outdir, filename+".odg")
    if args.verbosity > 0:
        print("Saving {}".format(path))
    template_object.save(path)
