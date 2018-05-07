import sys
import os
import csv
import toml
import xlrd
from PIL import Image, ImageFont, ImageDraw


def find_file(filename, style_dir=""):
    """Finding file in possible directories"""
    tmp_filename = os.path.join(INPUT_DIR, filename)
    if os.path.exists(tmp_filename):
        return tmp_filename
    tmp_filename = os.path.join(style_dir, filename)
    if os.path.exists(tmp_filename):
        return tmp_filename
    tmp_filename = os.path.join(os.path.dirname(
        os.path.dirname(os.getcwd())), filename)
    if os.path.exists(tmp_filename):
        return tmp_filename
    raise NameError("File not found")


def get_style(style_name):
    """Getting the style information"""
    style_filename = os.path.join("config", style_name, "layout.")
    if os.path.exists(style_filename + "toml"):
        style_filename += "toml"
    elif os.path.exists(style_filename + "ini"):
        style_filename += "ini"
    else:
        raise NameError("Layout file not found")
    return toml.load(style_filename)


def get_text_size(text, font, size, offset=0):
    """Getting the proper size in point of the text"""
    tmp_size = size[1]
    tmp_font = ImageFont.truetype(font, tmp_size)
    # while tmp_font.getsize(text)[0] <= size[0] and \
    #        tmp_font.getsize(tmp_text)[1] <= size[1] / (1 - offset):
    #    tmp_size += 1
    #    tmp_font = ImageFont.truetype(font, tmp_size)
    #tmp_size -= 1
    #tmp_font = ImageFont.truetype(font, tmp_size)
    #print (text, " ", tmp_size, " ", size)
    while tmp_font.getsize(text)[0] > size[0]:
        if tmp_size == 1:
            break
        tmp_size -= 1
        tmp_font = ImageFont.truetype(font, tmp_size)
    return [tmp_font, tmp_size]


def draw_text(s_img, text, font, color, pos, size, h_align="left", v_align="top", offset=0):
    """Drawing the text to the image"""
    new_img = s_img
    try:
        font_loc = find_file(font)
    except NameError:
        font_loc = font
    tmp_font = get_text_size(text, font_loc, size, offset)[0]
    text_size = [tmp_font.getsize(text)[0],
                 tmp_font.getsize(text)[1] * (1 - offset)]
    tmp_pos = [pos[0], pos[1] - tmp_font.getsize(text)[1] * offset]
    if h_align == "center":
        tmp_pos[0] += (size[0] - text_size[0]) / 2
    elif h_align == "right":
        tmp_pos[0] += size[0] - text_size[0]
    if v_align == "center":
        tmp_pos[1] += (size[1] - text_size[1]) / 2
    elif v_align == "bottom":
        tmp_pos[1] += size[1] - text_size[1]
    text += "　" #for Source Han Sans
    #if all(ord(c) < 128 for c in text): #for Source Han Sans
    #    tmp_pos[1] -= 2
    ImageDraw.Draw(new_img).text(tmp_pos, text, fill=color, font=tmp_font)
    #print (tmp_size," ",text_size," ", size[0],"*",size[1])
    return new_img


def generate_text(s_img, item_list, index, para_list):
    """Generating image of 'text' type"""
    #print(index, ",", para_list["position"])
    new_img = draw_text(s_img, item_list[index],
                        para_list["font"], para_list["color"], para_list["position"],
                        [para_list["width"], para_list["height"]],
                        para_list["horizontal_align"],
                        para_list["vertical_align"], para_list["offset"])
    return [new_img, index + 1]


def generate_colortext(s_img, item_list, index, para_list):
    """Generating image of 'colortext' type"""
    new_img = draw_text(s_img, item_list[index + 1],
                        para_list["font"], item_list[index], para_list["position"],
                        [para_list["width"], para_list["height"]],
                        para_list["horizontal_align"],
                        para_list["vertical_align"], para_list["offset"])
    return [new_img, index + 2]


def generate_vertitext(s_img, item_list, index, para_list):
    """Generating image of 'vertitext' type"""
    text = item_list[index]
    text_len = len(text)
    [tmp_font, tmp_size] = get_text_size(text[0], para_list["font"],
                                         [para_list["width"], int(para_list["height"] /
                                                                  text_len)],
                                         para_list["offset"])
    for i in text[1:]:
        tmp_fontset = get_text_size(i, para_list["font"],
                                    [para_list["width"], int(para_list["height"] /
                                                             text_len)],
                                    para_list["offset"])
        if tmp_fontset[1] >= tmp_size:
            [tmp_font, tmp_size] = tmp_fontset
    text_size = [0, 0]
    v_step = []
    for i in text:
        single_size = [tmp_font.getsize(i)[0],
                       tmp_font.getsize(i)[1] * (1 - para_list["offset"])]
        text_size[0] = max(text_size[0], single_size[0])
        text_size[1] += single_size[1]
        v_step.append(single_size[1])
        if i != text[-1]:
            text_size[1] += para_list["space"]
            v_step[-1] += para_list["space"]
    if para_list["vertical_align"] == "center":
        cur_v_pos = int((para_list["height"] - text_size[1]) / 2)
    elif para_list["vertical_align"] == "bottom":
        cur_v_pos = para_list["height"] - text_size[1]
    else:
        cur_v_pos = 0
    cur_v_pos += para_list["position"][1]
    for i in range(text_len):
        new_img = draw_text(s_img, text[i], para_list["font"], para_list["color"],
                            [para_list["position"][0], int(cur_v_pos)],
                            [para_list["width"], int(v_step[i])],
                            para_list["horizontal_align"], "center",
                            para_list["offset"])
        cur_v_pos += v_step[i]
    return [new_img, index + 1]


def generate_doubletext_nl(s_img, item_list, index, para_list):
    """Generating image of 'doubletext_nl' type"""
    single_height = int((para_list["height"] - para_list["space"])/2)
    new_img = draw_text(s_img, item_list[index],
                        para_list["font"], para_list["color"], para_list["position"],
                        [para_list["width"],
                         single_height], para_list["horizontal_align"],
                        para_list["vertical_align"], para_list["offset"])
    new_img = draw_text(s_img, item_list[index + 1],
                        para_list["font"], para_list["color"],
                        [para_list["position"][0], para_list["position"]
                         [1] + single_height + para_list["space"]],
                        [para_list["width"],
                         single_height], para_list["horizontal_align"],
                        para_list["vertical_align"], para_list["offset"])
    return [new_img, index + 2]


def generate_doubletext(s_img, item_list, index, para_list):
    """Generating image of 'doubletext' type"""
    if item_list[index + 1] == "":
        return [generate_text(s_img, item_list, index, para_list)[0], index + 2]
    i_size = [para_list["width"], para_list["height"]]
    [major_font, major_size] = get_text_size(item_list[index], para_list["font"],
                                             i_size, para_list["offset"])
    major_text_size = major_font.getsize(item_list[index])
    [minor_font, minor_size] = get_text_size(item_list[index + 1], para_list["font"],
                                             [i_size[0] - major_text_size[0] - para_list["space"],
                                              i_size[1]], para_list["offset"])
    while minor_size < para_list["minimum_point"] or \
            major_size - minor_size > para_list["maximum_diff"]:
        major_size -= 1
        major_font = ImageFont.truetype(para_list["font"], major_size)
        major_text_size = major_font.getsize(item_list[index])
        [minor_font, minor_size] = get_text_size(item_list[index + 1], para_list["font"],
                                                 [i_size[0] - major_text_size[0] -
                                                  para_list["space"], i_size[1]],
                                                 para_list["offset"])
    if major_size - minor_size <= para_list["minimum_diff"]:
        minor_size = max(para_list["minimum_point"],
                         major_size - para_list["minimum_diff"])
    minor_font = major_font = ImageFont.truetype(para_list["font"], minor_size)
    minor_text_size = minor_font.getsize(item_list[index + 1])
    new_img = draw_text(s_img, item_list[index],
                        para_list["font"], para_list["color"], para_list["position"],
                        [major_text_size[0], para_list["height"]], "left",
                        para_list["vertical_align"], para_list["offset"])
    minor_pos = para_list["position"].copy()
    minor_pos[0] += major_text_size[0] + para_list["space"]
    minor_textsize = [minor_text_size[0], para_list["height"]].copy()
    new_img2 = draw_text(new_img, item_list[index + 1],
                         para_list["font"], para_list["color"], minor_pos,
                         minor_textsize, para_list["horizontal_align_right"],
                         para_list["vertical_align"], para_list["offset"])
    #print (major_size,", ",minor_size,", ",item_list[index])
    return [new_img2, index + 2]


def generate_figure(s_img, item_list, index, para_list):
    """Generating image of 'figure' type"""
    try:
        if item_list[index] == "":
            return [s_img, index + 1]
    except KeyError:
        return [s_img, index + 1]
    #print(para_list, ", ", item_list[index])
    new_img = s_img
    try:
        imgfile = find_file(item_list[index])
    except NameError:
        imgfile = para_list["figure_alias"][item_list[index]]
    fig = Image.open(imgfile)
    if para_list["keep_aspect_ratio"] == 0:
        fig = fig.resize((para_list["width"], para_list["height"]))
        pos = para_list["position"]
    else:
        fig.thumbnail((para_list["width"], para_list["height"]))
        pos = para_list["position"].copy()
        if para_list["horizontal_align"] == "center":
            pos[0] += int((para_list["width"] - fig.size[0]) / 2)
        elif para_list["horizontal_align"] == "right":
            pos[0] += para_list["width"] - fig.size[0]
        if para_list["vertical_align"] == "center":
            pos[1] += int((para_list["height"] - fig.size[1]) / 2)
        elif para_list["vertical_align"] == "bottom":
            pos[1] += para_list["height"] - fig.size[1]
    new_img.paste(fig, pos, fig)
    return [new_img, index + 1]


# def generate_figuregroup(s_img, item_list, index, para_list):
#    """Generate image of 'figuregroup' type"""
#    new_img = s_img
#    return [new_img, index+1]


def process_section(style, layout_type, item_list):
    """Processing a single section"""
    switchfunc = {
        "text": generate_text,
        "colortext": generate_colortext,
        "vertitext": generate_vertitext,
        "doubletext_nl": generate_doubletext_nl,
        "doubletext": generate_doubletext,
        "figure": generate_figure
        #"figuregroup": generate_figuregroup
    }
    if layout_type == "figure_alias":
        raise NameError("Invalid type name")
    layout = style[layout_type]
    s_img = Image.open(find_file(layout["background"]))
    index = 0
    for i in layout["item"]:
        try:
            i_conf = layout["default_" + i["type"]].copy()
            i_conf.update(i)
        except KeyError:
            i_conf = i.copy()
        try:
            f_alias = style["figure_alias"]
            i_conf.update({"figure_alias": f_alias})
        except KeyError:
            pass
        [s_img, index] = switchfunc[i["type"]](
            s_img, item_list, index, i_conf)
    return s_img


def process_sheet(s_sheet):
    """Processing a single sheet"""
    s_index = 0
    s_len = len(s_sheet)
    while s_sheet[s_index][0] == "//":
        s_index += 1
        if s_index == s_len:
            raise NameError("No valid style")
    style_name = s_sheet[s_index][0]
    style = get_style(style_name)
    os.chdir(os.path.join("config", style_name))
    s_index += 1
    current_type = "default"
    main_img = None
    main_size = [0, 0]
    for i in s_sheet[s_index:]:
        if i[0] == "//":
            continue
        if i[0] != "":
            current_type = i[0]
        s_img = process_section(
            style, current_type, i[1:])
        s_size = list(s_img.size)
        if main_img is None:
            main_img = s_img
            main_size = s_size
        else:
            new_size = [max(main_size[0], s_size[0]), main_size[1] + s_size[1]]
            main_img = main_img.crop([0, 0] + new_size)
            new_box = [0, main_size[1], s_size[0], new_size[1]]
            main_img.paste(s_img, new_box)
            main_size = new_size
    os.chdir(os.path.dirname(os.path.dirname(os.getcwd())))
    return main_img


def read_csv_file(csv_filename):
    """Reading CSV file"""
    c_file = open(csv_filename, "r")
    content = list(csv.reader(c_file))
    c_file.close()
    return {"csvsheet": content}


def read_excel_file(excel_filename):
    """Reading .xlsx/.xls file"""
    x_book = xlrd.open_workbook(excel_filename)
    content = {}
    tmp_sheetname = x_book.sheet_names()
    for i in range(x_book.nsheets):
        tmp_content = []
        tmp_sheet = x_book.sheet_by_index(i)
        for j in range(tmp_sheet.nrows):
            tmp_content += [tmp_sheet.row_values(j)]
        content[tmp_sheetname[i]] = tmp_content
    return content


def valid_filename(input_filename):
    """Making filename valid"""
    return input_filename.translate(str.maketrans("*/\\<>:\"|", "--------"))


def main(argv=None):
    """Main function of HCTIWS"""
    global INPUT_DIR
    if argv is None:
        argv = sys.argv
    # display the version info
    print("HCTIWS Creates the Image with Sheets")
    print("       Made by ZMSOFT")
    print("version 2.67-3\n")
    # get input filename
    if len(argv) == 1:
        print("Usage: hctiws [INPUT_FILE [OUTPUT_DIRECTORY]]\n")
        input_filename = input("Input file (.csv, .xlsx, .xls): ")
        INPUT_DIR = os.path.dirname(input_filename)
    elif len(argv) == 2:
        input_filename = argv[1]
        INPUT_DIR = os.path.dirname(input_filename)
    else:
        input_filename = argv[1]
        INPUT_DIR = argv[2]
    # open worksheet/sheet file
    if input_filename[-4:] == ".csv":
        content = read_csv_file(input_filename)
    elif input_filename[-5:] == ".xlsx" or input_filename[-4:] == ".xls":
        content = read_excel_file(input_filename)
    # process and save the result of each sheet
    for i in content:
        tmp_img = process_sheet(content[i])
        # tmp_img.show()  # show the image before saving for debug
        tmp_name = os.path.join(
            INPUT_DIR, valid_filename(
                os.path.basename(input_filename)).
            rsplit(".", 1)[0] + "_" + valid_filename(i))
        if os.path.exists(tmp_name + ".png"):
            for j in range(1, 100):
                if not os.path.exists(tmp_name + "-" + str(j) + ".png"):
                    tmp_name += "-" + str(j) + ".png"
                    break
            else:
                raise NameError("Too many duplicated file names")
        else:
            tmp_name += ".png"
        tmp_img.save(tmp_name)
        print(tmp_name, "DONE")
    return 0


INPUT_DIR = ""
if __name__ == "__main__":
    sys.exit(main())
