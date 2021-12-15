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
        print("Done!")


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
        print("Done!")

def tex_op(source, output):   #TODO: 添加ebook
    return

def ebook_json(data, ebook):
    man_json(data, ebook)


def to_json(data, man, op):
    data = xlrd.open_workbook(data)
    man_json(data, man)
    op_json(data, op)
    ebook_json(data, op)


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
    to_json("data/data.xls", 'data/man.json', 'data/op.json')
    # to_json("data/data.xls", 'data/man.json', 'data/op.json', 'data/ebook.json')//TODO:添加ebook
    print("DONE: to json")
    tex_man("data/man.json", "tex/man/man.tex")
    tex_op("data/op.json", "tex/op/op.tex")
    # tex_ebook("data/ebook.json", "tex/ebook/ebook.tex")//TODO:添加ebook
    compile_tex(root, "man")
    compile_tex(root, "op")
    print("All done!!!")
