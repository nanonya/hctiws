# HCTIWS 使用说明 #

本工具可以从表格生成固定格式的图片。

    hctiws [INPUT_FILE [OUTPUT_DIRECTORY]]

如不指定参数，则程序会要求手动输入表格的文件名称，并在表格所在的目录下生成图片。

如果指定输入文件作为参数，默认直接在表格所在的目录下生成图片，但可选手动指定输出目录。

本工具使用 Python 3 开发。依赖的第三方库包括 pillow、xlrd、toml。
请使用方便的方式如 pip 安装。

这一工具是为了特定目的开发和使用的，因此未来可能会较少得到更新，请理解。

## 表格 ##
表格文件格式支持 csv、xlsx、xls 三种格式。csv 格式需要以逗号分隔。未来会加入 ods 格式的支持。

表格中，A 列为配置信息，具体规则如下：

- 任何以 // 开头的行都会被忽略，其后的内容算作注释。
- 第一个有效的行的 A 列内容为所使用的生成配置 / 风格。具体来说，会以其内容为目录名称在 config 目录下寻找。找到后，按该文件夹内的 layout.toml 或 layout.ini 作为生成图片时的配置文件，以此来安排项目的布局。该行后面的内容都会被忽略。
- 在此之后，所有有效的行，如果 A 列具有内容，则以此内容为名称在配置文件内寻找局部布局类型；如果 A 列为空，则使用前一行的布局类型。其后的内容，会按照该布局类型的规则被添加到布局类型指定的背景图中，最终这张图片会被附加在前面生成的图片底部。

## 配置文件 ##

本工具使用 toml 作为每种生成配置 / 风格的文件格式。为方便在部分 Windows 系统上打开，允许使用 ini 作为扩展名。配置文件名为 layout.toml 或 layout.ini，置于 config/[配置 / 风格名称] 目录下。

配置文件的格式，以下列内容为例：

    [content1]
        background="example.png"
        [content1.default_text]
            color="black"
            font="example_font.otf"
            height=50
        [[content1.item]]
            type="text"
            color="grey"
            position=[10,20]
            width=500
            horizontal_align="left"
            vertical_align="top"
            minimum_point=10
            offset=0
        [[content1.item]]
            type="figure"
            keep_aspect_ratio=1
            horizontal_align="center"
            vertical_align="center"
            position=[100,20]
            width=40
            height=40
        [[content1.item]]
            ...
    [content2]
        ...
    [figure_alias]
        "circle"="circle.png"

其中，content1 和 content2 为局部布局类型名称，对应表格 A 列内容中除第一个之外的有效内容所指定的项目。

每一种局部布局可以指定包含若干个内容部件。例如，对于 content1 布局类型，一行表格可以如下指定：

    content1    txt0    fig0    ...

（如果表格的前一行亦为 content1 类型，首列内容可为空。）

生成的对应局部图片，其首个 text 项目填入文字 txt，而首个 figure 项目会寻找图片 fig0 填入。

内容部件的类型包括：

- text 文本。

- colortext 彩色文本。需要两个表格项，使用一个参数指定颜色（可以为颜色名称如 black 或 Web 颜色代码），后一个参数指定文本内容

- vertitext 竖排文本。

- doubletext 左右双文本。需要两个表格项。

- doubletext_nl 上下双文本。需要两个表格项。

- figure 图形。首先在表格所在的文件夹内寻找是否有指定的文件名，其次在风格文件夹内寻找，再其次在本工具所在文件夹内寻找。如果均无法找到，则查询配置文件内的 figure_alias 表。该表内以类似 "fig1"="figure1.png" 的方式指定图像别名。对应的图像亦以上述优先级寻找，但不会二次查找别名。

各部件（item）均包含多个属性。其中，type 属性指定类型，其余的属性种类因类型而定。配置中表名为「default_」加类型的表中的内容会被作为该类型的默认设定，但优先级低于单独的设定。

## text 类型属性

- font 指定字体，文字类部件使用。须指定字体文件，以上述优先级寻找，但不会二次查找别名。可使用系统字体。

- color 指定颜色。内容可以为颜色名称如 black 或 Web 颜色代码。

- position 在该局部布局内的左上角位置。

- width 宽度。

- height 高度。

- minimum_point 最小字号。

- horizontal_align 横向位置，可选 left、center 或 right。

- vertical_align 纵向位置，可选 top、center 或 bottom。

- offset 偏移，其值不小于 0，不大于 1。为了应对部分中文字体的留空问题，将上部该比例的部分自动视作不存在。

## colortext 类型属性

与 text 类型一致，但不包含 color，因该部分为用户指定。

## vertitext 类型属性

与 text 类型一致，但不包含 minimum_point，新增以下类型：

- space 指定竖排时两字之间的空隙。

## doubletext 类型属性

与 text 类型一致，但新增以下类型：

- horizontal_align_right 右侧文字的横向位置。相应的，horizontal_align 只控制左侧文字。

- minimum_diff 两段文字最小字号差异。

- maximum_diff 两段文字最大字号差异。

- prior 指定左右两边何者字号更大，0 为左方，1 为右方。

- space 指定两部分文字间距。

## doubletext_nl 类型属性

与 text 类型一致，但新增以下类型：

- horizontal_align_down 下方文字的纵向位置。相应的，horizontal_align 只控制上方文字。

- space 指定两部分文字间距。

## figure 类型属性

- keep_aspect_ratio 保持宽高比，0 为不保持，1 为保持。

- position 在该局部布局内的左上角位置。

- width 宽度。

- height 高度。

- horizontal_align 横向位置，可选 left、center 或 right。

- vertical_align 纵向位置，可选 top、center 或 bottom。
