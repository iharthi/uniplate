import zipfile
import odf.table
import odf.opendocument
import odf.text
import os.path


def cell_value(cell):
    """Returns a cell as text value"""
    ps = cell.getElementsByType(odf.text.P)
    text_content = ""
    for p in ps:
        text_content += node_value(p)
    return text_content


def node_value(node):
    """Tries to convert an Element into text value"""
    text = ''
    if hasattr(node, "childNodes"):
        for n in node.childNodes:
            text += node_value(n)
    if hasattr(node, 'data'):
        text += node.data
    return text


class BaseTableLoader(object):
    @classmethod
    def load_table(cls, args):
        raise NotImplementedError()


class TableLoader(BaseTableLoader):
    @classmethod
    def load_table(cls, args):
        table_object = None

        try:
            table_object = odf.opendocument.load(args.table)
        except zipfile.BadZipFile:
            print(
                "{} is not a valid zip archive (which means it's also not a .ods document for sure)".format(args.table))
            exit()

        if table_object is None or table_object.mimetype != 'application/vnd.oasis.opendocument.spreadsheet':
            print("{} is not an Open Document Chart".format(args.table))
            exit()

        naming = args.naming

        uniplate_globals = {}

        for sheet in table_object.spreadsheet.getElementsByType(odf.table.Table):
            s_name = sheet.getAttribute("name")
            if s_name == 'uniplate_globals':
                rows = sheet.getElementsByType(odf.table.TableRow)
                if len(rows) < 2:
                    continue
                header_row = rows[0]
                data_row = rows[1]
                global_header = [cell_value(cell) for cell in header_row.getElementsByType(odf.table.TableCell)]

                cells = []

                for cell in data_row.getElementsByType(odf.table.TableCell):
                    value = cell_value(cell)
                    columns = 1
                    for attr in cell.attributes:
                        if 'number-columns-repeated' in attr:
                            columns = int(cell.attributes[attr])
                    while len(cells) < len(global_header) and columns > 0:
                        cells.append(value)
                        columns-=1
                for i in range(len(global_header)):
                    try:
                        value = cells[i]
                    except IndexError:
                        value = ""
                    cls.process_cell(uniplate_globals, global_header[i], value, args)

        table = []
        for sheet in table_object.spreadsheet.getElementsByType(odf.table.Table):
            s_name = sheet.getAttribute("name")
            if s_name == 'uniplate_globals':
                continue
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

                for key in uniplate_globals:
                    row_dictionary[key] = uniplate_globals[key]

                cells = []

                for cell in row.getElementsByType(odf.table.TableCell):
                    value = cell_value(cell)
                    columns = 1
                    for attr in cell.attributes:
                        if 'number-columns-repeated' in attr:
                            columns = int(cell.attributes[attr])
                    while len(cells) < len(header) and columns > 0:
                        cells.append(value)
                        columns-=1

                for i in range(len(header)):
                    try:
                        value = cells[i]
                    except IndexError:
                        value = ""
                    cls.process_cell(row_dictionary, header[i], value, args)
                if args.verbosity > 2:
                    print(row_dictionary)
                cls.post_process_row(row_dictionary)
                table.append(row_dictionary)
        cls.post_process_table(table)
        return table, naming

    @classmethod
    def process_cell(cls, row_dictionary, header, value, args):
        if '::' in header:
            # Key-value column
            try:
                name, key = header.split('::')
                if name not in row_dictionary:
                    row_dictionary[name] = []
                if not args.skip_empty or value != "":
                    row_dictionary[name].append((key, value))
                return # Prevent processing as a regular column
            except ValueError:
                pass  # Will process as a regular column
        # Regular column
        row_dictionary[header] = value

    @classmethod
    def post_process_row(cls, row_dictionary):
        pass

    @classmethod
    def post_process_table(cls, table):
        pass


class BaseTemplater(object):
    def template_file(self, file):
        raise NotImplementedError()


class Templater(BaseTemplater):
    def __init__(self, args, naming):
        """Tries to load a template, exits if the template is corrupted or not existing"""
        self.args = args
        self.naming = naming
        self.save_callback = self.save

    def reload_template(self):
        my_template_object = None
        try:
            my_template_object = odf.opendocument.load(self.args.template)
        except zipfile.BadZipFile:
            print(
                "{} is not a valid zip archive (which means it's also not a .ods document for sure)".format(
                    self.args.template))
            exit()

        if my_template_object is None or my_template_object.mimetype != 'application/vnd.oasis.opendocument.graphics':
            print("{} is not an Open Document Chart".format(self.args.template))
            exit()

        self.template_object = my_template_object

    def template_string(self, string, name, text):
        """Substitutes {name} to text in the string"""
        if self.args.verbosity > 4:
            print(string)
        if isinstance(text, str):
            return string.replace("{" + name + "}", text)
        elif isinstance(text, list):
            for i in range(len(text)):
                string = string.replace("{" + name + "::key::" + str(i) + "}", text[i][0])
                string = string.replace("{" + name + "::value::" + str(i) + "}", text[i][1])
            while "{" + name + "::key::" in string:
                start = string.index("{" + name + "::key::")
                end = string.index("}", start) + 1
                if self.args.fill_with_last:
                    string = string[:start] + text[-1][0] + string[end:]
                else:
                    string = string[:start] + string[end:]
            while "{" + name + "::value::" in string:
                start = string.index("{" + name + "::value::")
                end = string.index("}", start) + 1
                if self.args.fill_with_last:
                    string = string[:start] + text[-1][1] + string[end:]
                else:
                    string = string[:start] + string[end:]
        return string

    def template_node(self, node, name, text):
        """Substitutes {name} to text within an Element and children.
         {name} has to be within a single Element."""
        if hasattr(node, "childNodes"):
            for n in node.childNodes:
                self.template_node(n, name, text)
        if hasattr(node, 'data'):
            if isinstance(text, str) and "{" + name + "}" in node.data or \
                            isinstance(text, list) and (
                                    "{" + name + "::key::" in node.data or "{" + name + "::value::"):
                node.data = self.template_string(node.data, name, text)

    def template(self, node, name, text):
        """Substitutes {name} to text within a Document"""
        ps = node.getElementsByType(odf.text.P)
        for p in ps:
            self.template_node(p, name, text)

    def template_file(self, row):
        """Processes single file with template engine"""
        self.reload_template()

        self.preprocess_file(row)

        filename = self.naming
        for key in row:
            self.template(self.template_object, key, row[key])
            filename = self.template_string(filename, key, row[key])

        self.postprocess_file(row)

        if os.path.isfile(os.path.join(self.args.outdir, filename + ".odg")):
            suffix = 1
            while os.path.isfile(os.path.join(self.args.outdir, filename + "_{}".format(suffix) + ".odg")):
                suffix += 1
            filename += "_{}".format(suffix)
        path = os.path.join(self.args.outdir, filename + ".odg")
        if self.args.verbosity > 0:
            print("Saving {}".format(path))

        self.save_callback(path)

    def preprocess_file(self, row):
        pass

    def postprocess_file(self, row):
        pass

    def save(self, path):
        self.template_object.save(path)