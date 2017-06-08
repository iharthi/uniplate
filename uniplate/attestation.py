import uniplate_engine
import datetime
import locale
import odf.draw

locale.setlocale(locale.LC_ALL, '')


class TableLoader(uniplate_engine.TableLoader):
    prefix = '::'
    marks = {
        '3': "3 (удовлетворительно)",
        '4': "4 (хорошо)",
        '5': "5 (отлично)",
    }

    @classmethod
    def process_cell(cls, row_dictionary, header, value, args):
        if header.startswith(cls.prefix):
            row_dictionary[header[len(cls.prefix):].lower()] = value
        elif header.strip() != "":
            value = cls.format_mark_value(value)
            if 'mark' not in row_dictionary:
                row_dictionary['mark'] = []
            if not args.skip_empty or value != "":
                row_dictionary['mark'].append((header, value))

    @classmethod
    def post_process_row(cls, row_dictionary):
        try:
            row_dictionary['name'] = " ".join(
                [row_dictionary['name1'],row_dictionary['name2'],row_dictionary['name3']])
        except KeyError:
            row_dictionary['name'] = ""

        try:
            row_dictionary['birth'] = datetime.date(
                year=int(row_dictionary["birthyear"])+
                    ((int(str(datetime.datetime.now().year-50)[:2])+0)*100 if
                        int(row_dictionary["birthyear"]) > int(str(datetime.datetime.now().year-50)[-2:]) else
                        (int(str(datetime.datetime.now().year - 50)[:2])+1) * 100),
                month=int(row_dictionary["birthmonth"]),
                day=int(row_dictionary["birthday"])
            ).strftime('%d %B %Y года')
        except (ValueError, KeyError, OverflowError) as e:
            row_dictionary['birth'] = ""
        row_dictionary['mark'].append(("Z", "Z"))

    @classmethod
    def post_process_table(cls, table):
        empty = []
        for i in range(len(table)):
            if 'name' not in table[i] or table[i]['name'].strip() == '':
                empty.insert(0,i)
        for i in empty:
            del table[i]

    @classmethod
    def format_mark_value(cls, mark):
        if mark in cls.marks:
            return cls.marks[mark]
        else:
            return ""


class Templater(uniplate_engine.Templater):
    def preprocess_file(self, row):
        if 'mark' in row:
            if len(row['mark']) - 1 > 21:
                # Second page active
                # Remove the big Z
                for line in self.template_object.getElementsByType(odf.draw.Line):
                    if 'bigZ_' in line.getAttribute('name'):
                        line.parentNode.removeChild(line)
            else:
                # Second page inactive
                # Remove the mark placeholders
                erase = []
                for i in range(21,42):
                    erase.append("mark::key::"+str(i)+"")
                    erase.append("mark::value::" + str(i) + "")
                for key in erase:
                    self.template(self.template_object, key, "")

