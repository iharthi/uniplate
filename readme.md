# General purpose (universal) template engine for OpenDocument files

## Requirements

`uniplate.py` is a standalone script, no other files are required to run it, and the only required python module
is `odfpy`. Tested with `odfpy-1.3.3`. Developed with `python3`, but works both with `python2`
and `python3`.


## Installation

* Step 1: get the script

    `git clone https://github.com/iharthi/uniplate.git uniplate`

* Step 2: install required module(s)

    `pip install -r uniplate/requirements.txt `


## Basic usage

`uniplate.py template.odg table.ods`

The script takes all sheets in the table spreadsheet file and interprets first row of each sheet as a header containing
field names. It will then create a copy of template file for each row, and substitute any template field marks with the
respective value for the row. The template field marks consist of field name in {} brackets, like `{field}`.


## More complex usage

`uniplate.py [-h] [-o OUTDIR] [-v] [-n NAMING] [-s SHEET] template table`

positional arguments:

 * `template` — Template file to use (should work with any open document, tested on .odg)
 * `table` — Spreadsheet file with data (.ods file)

optional arguments:
 * `-h`, `--help` — show this help message and exit
 * `-o OUTDIR`, `--outdir OUTDIR` — Output directory
 * `-v`, `--verbosity` — Be verbose
 * `-n NAMING`, `--naming NAMING` — File naming pattern, like {field1}_{field2}. If not set, first field name of first
  sheet in the table file is used. The script will never overwrite files, numeric suffix is added instead.
 * `-s SHEET`, `--sheet SHEET` — Sheet to process (can occur more than once). If no occurrences, all sheets are
  processed


## Known bugs

It is possible to create a template with template field mark split among several OpenDocument Elements. Even if the
 Elements have same formatting and look visually identical, the text inside is split, interrupting the field mark
 sequence and preventing it from being replaced with field value.

There is no easy programmatic workaround to this problem, but it is easy to fix the template by hand, by erasing
the whole template field mark and retyping in from scratch.
