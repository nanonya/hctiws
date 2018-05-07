# HCTIWS 使用说明 #

本工具可以从表格生成固定格式的图片。

    hctiws [INPUT_FILE [OUTPUT_DIRECTORY]]

如不指定参数，则程序会要求手动输入表格的文件名称，并在表格所在的目录下生成图片。

如果指定输入文件作为参数，默认直接在表格所在的目录下生成图片，但可选手动指定输出目录。

## 表格 ##
表格文件格式支持 csv、xlsx、xls 三种格式。csv 格式需要以逗号分隔。未来会加入 ods 格式的支持。

表格中，A 列为配置信息，具体规则如下：

- 任何以 // 开头的行都会被忽略，其后的内容算作注释。
- 第一个有效的行的 A 列内容为所使用的生成配置 / 风格。具体来说，会以其内容为目录名称在 config 目录下寻找。找到后，按该文件夹内的 layout.toml 或 layout.ini 作为生成图片时的配置文件，以此来安排项目的布局。该行后面的内容都会被忽略。
- 在此之后，所有有效的行，如果 A 列具有内容，则以此内容为名称在配置文件内寻找布局类型；如果 A 列为空，则使用前一行的布局类型。其后的内容，会按照该布局类型的规则被添加到布局类型指定的背景图中，最终这张图片会被附加在前面生成的图片底部。

## 配置文件 ##

本工具使用 toml 作为每种生成配置 / 风格的文件格式。为方便在部分 Windows 系统上打开，允许使用 ini 作为扩展名。配置文件名为 layout.toml 或 layout.ini，置于 config/[配置 / 风格名称] 目录下。

配置文件的格式，以下列内容为例：

    [content1]
        background="example.png"
        [content1.default_text]
            color="black"
            font="example_font.otf"
            height=50
        [content1.item]
            type="text"
            color="grey"
            position=[10,20]
            width=500
            horizontal_align="left"
            vertical_align="top"
            minimum_point=10
            offset=0
        [[content.item]]
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

其中，content1 和 content2 为布局名称，即表格 A 列内容中除第一个之外的有效内容所指定的项目。
