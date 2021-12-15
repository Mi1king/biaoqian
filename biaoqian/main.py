import os
import json
import xlrd


def check_length(line):
    ch_counter = 0
    en_counter = 0
    for ch in line:
        if u'\u4e00' <= ch <= u'\u9fff':
            ch_counter += 1
        else:
            en_counter += 1
    length = ch_counter * 2 + en_counter
    return length <= 85


def tex_man(source, output):
    with open(source, "r", encoding='utf8') as f:
        data = json.load(f)
        # print(data)
        file = open("tex/document.txt", "r+", encoding='utf8')
        lines = file.readlines()
        l1 = '\\begin{center}\n'
        l2 = '\\chuhao{\\textbf{'
        l3 = '}\\\\}\n'
        l4 = '\\vspace{0.8cm}\n'
        l5 = '\\yihao{\\textbf{'
        l6 = '}\\\\}\n'
        l7 = '\\vspace{0.8cm}\n'
        l8 = '\\yihao{'
        l9 = '\\\\}\n'
        l10 = '\\vspace{0.4cm}\n'
        l11 = '\erhao{'
        l12 = '}\n'
        l13 = '\\vspace{6cm}\n'
        ln = '\\end{center}\n'
        lnp = '\\newpage\n'
        lf = '\\end{document}'
        # # print(line, "a")
        file1 = open(output, 'w', encoding='utf8')
        count = 0
        max_length = 32
        for d in data:
            flag = check_length(data[d][2])
            if "&" in str(data[d][2]):
                data[d][2] = str(data[d][2]).replace("&", "\\&")
            if "$" in str(data[d][2]):
                data[d][2] = str(data[d][2]).replace("$", "\\$")
            if "_" in str(data[d][2]):
                data[d][2] = str(data[d][2]).replace("_", "\\_")
            lines.append(l1)
            lines.append(l2 + d + l3)
            lines.append(l4)
            lines.append(l5 + data[d][1] + l6)
            lines.append(l7)
            if data[d][0] == "":
                lines.append('\\vspace{1cm}\n')
            else:
                lines.append(l8 + data[d][0] + l9)
                lines.append(l10)
            if count == 0:
                lines.append(l11 + data[d][2] + l12)
                lines.append(ln)
                lines.append(l13)
                count = 1
                if not flag:
                    lines.append(lnp)
                    count = 0
            else:
                if not flag:
                    lines.append(lnp)
                lines.append(l11 + data[d][2] + l12)
                lines.append(ln)
                lines.append(lnp)
                count = 0
        lines.append(lf)
        file.close()
        file1.writelines(lines)
        file1.close()
        f.close()


def tex_op(source, output):
    with open(source, "r", encoding='utf8') as f:
        data = json.load(f)
        # print(data)
        file = open("tex/document.txt", "r+", encoding='utf8')
        lines = file.readlines()
        l1 = '\\begin{center}\n'
        l2 = '\\chuhao{\\textbf{'
        l3 = '}\\\\}\n'
        l4 = '\\vspace{0.8cm}\n'
        l5 = '\\yihao{\\textbf{'
        l6 = '}\\\\}\n'
        l7 = '\\vspace{0.8cm}\n'
        l8 = '\\yihao{'
        l9 = '\\\\}\n'
        l10 = '\\vspace{0.2cm}\n'
        l11 = '\erhao{'
        l12 = '\\\\}\n'
        l13 = '\\vspace{4cm}\n'
        ln = '\\end{center}\n'
        lnp = '\\newpage\n'
        lf = '\\end{document}'
        # # print(line, "a")
        file1 = open(output, 'w', encoding='utf8')
        count = 0
        max_length = 32
        for d in data:
            flag = check_length(data[d][2])
            if "&" in str(data[d][2]):
                data[d][2] = str(data[d][2]).replace("&", "\\&")
            if "$" in str(data[d][2]):
                data[d][2] = str(data[d][2]).replace("$", "\\$")
            if "_" in str(data[d][2]):
                data[d][2] = str(data[d][2]).replace("_", "\\_")
            if "^" in str(data[d][2]):
                data[d][2] = str(data[d][2]).replace("^", "\\^")
            lines.append(l1)
            lines.append(l2 + d + l3)
            lines.append(l4)
            lines.append(l5 + data[d][1] + l6)
            lines.append(l7)
            if data[d][0] == "":
                lines.append('\\vspace{1cm}\n')
            else:
                lines.append(l8 + data[d][0] + l9)
                lines.append(l10)
            lines.append(l11 + "RMB " + data[d][3] + l12)
            lines.append(l11 + data[d][2] + l12)
            if not flag:
                lines.append(lnp)
            lines.append(ln)
            if count == 0:
                lines.append(l13)
                count = 1
            else:
                lines.append(lnp)
                count = 0
        lines.append(lf)
        file.close()
        file1.writelines(lines)
        file1.close()
        f.close()


def tex_ebook(source, output):  # TODO: 添加ebook
    with open(source, "r", encoding='utf8') as f:
        data = json.load(f)
        # print(data)
        file = open("tex/document.txt", "r+", encoding='utf8')
        lines = file.readlines()
        l1 = '\\begin{center}\n'
        l2 = '\\chuhao{\\textbf{'
        l3 = '}\\\\}\n'
        l4 = '\\vspace{0.8cm}\n'
        l5 = '\\yihao{\\textbf{'
        l6 = '}\\\\}\n'
        l7 = '\\vspace{0.8cm}\n'
        l8 = '\\yihao{'
        l9 = '\\\\}\n'
        l10 = '\\vspace{0.4cm}\n'
        l11 = '\erhao{'
        l12 = '}\n'
        l13 = '\\vspace{6cm}\n'
        ln = '\\end{center}\n'
        lnp = '\\newpage\n'
        lf = '\\end{document}'
        le = '\\textcolor{red}{text(E-BOOK)\\\\电子书由任课老师直接发放}'

        # # print(line, "a")
        file1 = open(output, 'w', encoding='utf8')
        count = 0
        max_length = 32
        for d in data:
            flag = check_length(data[d][2])
            if "&" in str(data[d][2]):
                data[d][2] = str(data[d][2]).replace("&", "\\&")
            if "$" in str(data[d][2]):
                data[d][2] = str(data[d][2]).replace("$", "\\$")
            if "_" in str(data[d][2]):
                data[d][2] = str(data[d][2]).replace("_", "\\_")
            lines.append(l1)
            lines.append(l2 + '\\textcolor{red}{' + d + '}' + l3)
            lines.append(l4)
            lines.append(l5 + data[d][1] + l6)
            lines.append(l7)
            if data[d][0] == "":
                lines.append('\\vspace{1cm}\n')
            else:
                lines.append(l8 + data[d][0] + l9)
                lines.append(l10)
            if count == 0:
                lines.append(l11 + data[d][2] + l12)
                lines.append(le)
                lines.append(ln)
                lines.append(l13)
                count = 1
                if not flag:
                    lines.append(lnp)
                    count = 0
            else:
                if not flag:
                    lines.append(lnp)
                lines.append(l11 + data[d][2] + l12)
                lines.append(ln)
                lines.append(lnp)
                count = 0
        lines.append(lf)
        file.close()
        file1.writelines(lines)
        file1.close()
        f.close()
    return


def sheet_to_json(file_name, json_file_name, sheet_name):
    wb = xlrd.open_workbook(file_name)
    table = wb.sheet_by_name(sheet_name)
    book_dic = {}
    n_rows = table.nrows
    for r in range(1, n_rows - 1):
        isbn = str(table.cell_value(r, 2))
        book_name = str(table.cell_value(r, 3))
        code = str(table.cell_value(r, 1))
        no = str(table.cell_value(r, 0))
        if sheet_name == 'Optional':
            price = str(table.cell_value(r, 4))
            book_dic[no] = [isbn, code, book_name, price]
        else:
            book_dic[no] = [isbn, code, book_name]

    with open(json_file_name, 'w') as f:
        json.dump(book_dic, f, indent=4)


# def to_json(file_name, json_file_name,sheet_name):
#     wb = xlrd.open_workbook(file_name)
#     sheet_to_json(wb, json_file_name, sheet_name)

# def to_json(data, man='', op='', ebook=''):
#     data = xlrd.open_workbook(data)
#     if man != '':
#         man_json(data, man)
#     if op != '':
#         op_json(data, op)
#     if ebook != '':
#         sheet_to_json(data, ebook, 'Ebook')


def man_json(data, man):
    table = data.sheet_by_index(0)
    book_man = {}
    n_rows = table.nrows
    for r in range(1, n_rows - 1):
        isbn = str(table.cell_value(r, 2))[:-2]
        book_name = str(table.cell_value(r, 3))
        code = str(table.cell_value(r, 1))
        no = str(table.cell_value(r, 0))[:-2]
        book_man[no] = [isbn, code, book_name]

    with open(man, 'w') as f:
        json.dump(book_man, f, indent=4)


def op_json(data, op):
    table = data.sheet_by_index(1)
    book_op = {}
    n_rows = table.nrows
    for r in range(1, n_rows - 1):
        isbn = str(table.cell_value(r, 2))[:-2]
        book_name = str(table.cell_value(r, 3))
        code = str(table.cell_value(r, 1))
        no = str(table.cell_value(r, 0))
        price = str(table.cell_value(r, 4))
        book_op[no] = [isbn, code, book_name, price]

    with open(op, 'w') as f:
        json.dump(book_op, f, indent=4)


def compile_tex(path, filename):
    tex_path = os.path.join(path, "tex", filename, filename)
    output_path = os.path.join(path, "output")
    command = "xelatex " + "-output-directory=" + output_path + " " + tex_path + ".tex"
    os.system(command)
    file_full_path = os.path.join(output_path, filename) + ".aux"
    if os.path.exists(file_full_path):
        os.remove(file_full_path)
    file_full_path = os.path.join(output_path, filename) + ".log"
    if os.path.exists(file_full_path):
        os.remove(file_full_path)


if __name__ == '__main__':
    root = os.getcwd()
    output_dir = 'output'
    # json_folder = os.path.join(output_dir, 'json')

    file_name = "data/Textbook Information AY2021-22 S2 WEEK14-做标签(useful_col).xls"
    mandatory_sheet_name = 'Mandatory'
    optional_sheet_name = 'Optional'
    ebook_sheet_name = 'Ebook'
    # mandatory_json_name = os.path.join(json_folder, 'man.json')
    # optional_json_name = os.path.join(json_folder, 'op.json')
    # ebook_json_name = os.path.join(json_folder, 'ebook.json')
    mandatory_json_name = 'data/man.json'
    optional_json_name = 'data/op.json'
    ebook_json_name = 'data/ebook.json'

    # Mandatory
    sheet_to_json(file_name, mandatory_json_name, mandatory_sheet_name)
    print("INFO: json file generation for Mandatory")
    tex_man(mandatory_json_name, "tex/man/man.tex")
    print("INFO: tex file generation for Mandatory")
    # compile_tex(root, "man")
    # print("DONE: pdf generation for Mandatory")

    # Optional
    sheet_to_json(file_name, optional_json_name, optional_sheet_name)
    print("INFO: json file generation for Optional")
    tex_op(optional_json_name, "tex/op/op.tex")
    print("INFO: tex file generation for Optional")
    # compile_tex(root, "op")
    # print("DONE: pdf generation for Optional")

    # Ebook
    sheet_to_json(file_name, ebook_json_name, ebook_sheet_name)
    print("INFO: json file generation for Ebook")
    tex_ebook(ebook_json_name, "tex/ebook/ebook.tex")
    print("INFO: tex file generation for Ebook")
    compile_tex(root, "ebook")
    print("DONE: compile tex for Ebook")

    print("All done!!!")
