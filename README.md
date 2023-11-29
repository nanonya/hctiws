# hctiws
[简体字](README%20(zh_CN).md)

HCTIWS is a tool to generate images in configured format from spreadsheet files.

By default, this tool saves the images to the directory of the spreadsheet file.
If no parameter is given, the program requires the user to input the file name manually.

Usage:
`hctiws [INPUT_FILE [OUTPUT_DIRECTORY]]`

This tool is developed with Python 3 and requires third-party libraries including pillow, xlrd and toml. You can use `pip` to install them.

HCTIWS is developed for a specific purpose and a specific user. Thus, please understand that there will be few updates in the future.

## Format of spreadsheet
Three formats (.csv, .xlsx, .xls) are supported as input. The separator of CSV files must be the comma character.

In the sheets, column A is associated to the configuration.
The rules are:
- Any row begins with `//` is ignored, and the whole row is considered as comments.
- The content in the cell A of first valid row specifies the name of configuration / style used in this session. To be exact, the program tries to find it in `config` directory, and the file `config/[NAME]/layout.toml` or `config/[NAME]/layout.ini` defines the style and layout of current session.
- Subsequent rows are regarded as contents of target image. The cell A decides which "partial layout" is used for placing contents in this row; if the cell A is empty, the previoud partial layout is reused. Other cells in the row stand for the contents filled in this part of image. All parts of the image are joined vertically.


## Configurations
HCTIWS uses TOML as the format of each configuration / style. It also accepts `.ini` extension as its TOML configuration file to make it more convenient to open the file in Windows operating system. The file name should be `layout.toml` or `layout.ini` and placed in `config/[NAME]/` directory.

An example of valid configuration is as follows:
```
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
```

In this example, `content1` and `content2` are names of "partial layout" types.
Each partial layout contains some items. For "content1" layout type, a row `content1,txt0,fig0,...` means:
- "content1" layout type is used.
- The first item's content is "txt0", the second item's content is "fig0".
- According to the config above, the first item is a "text" item, and the second one is a "figure" item; then the text "txt0" and image from the file "fig0" are placed on this part.

## Types of items
The supported item types are:
- `text`: one cell for the text content.
- `colortext`: require 2 cells, the first cell is color (in color name or HTML color codes), the latter contains text contents.
- `vertitext`: vertical text.
- `doubletext`: text from 2 cells in a row.
- `doubletext_nl`: require 2 cells, 2 lines of text.
- `figure`: find images in target directory, `config/[NAME]/` directory, and then the directory contains HCTIWS. It supports aliases set in `figure_alias` from the configuration.

Each item type has multiple parameters. Default parameters can be set in `default_[type]`.

## Parameters of item types
### `text`
- `font`: the file name of font file (support system fonts or custom fonts)
- `color`: supports color names (like "`black`") or HTML color codes
- `position`
- `width`
- `height`
- `minimum_point`: minimum size
- `horizontal_align`: valid values are "`left`", "`center`" and "`right`"
- `vertical_align`: valid values are "`top`", "`center`" and "`bottom`"
- `offset`: vertical offset of text; its value should be float number between 0 and 1

### `colortext`
The same as `text`, without the `color` parameter.

### `vertitext`
Nearly the same as `text`, without the `color` parameter.
The `vertitext` type has a new parameter:
- `space`: controls the space of each two characters

### `doubletext`
Nearly the same as `text` with some new parameters:
- `horizontal_align_right` (`horizontal_align` controls the left side)
- `minimum_diff`: minimum difference of sizes of both sides
- `maximum_diff`: maximum difference of sizes of both sides
- `prior`: controls which side can be larger - `0` for the left side, `1` for the right side
- `space`: controls the space of two sides of text

### `doubletext_nl`
Nearly the same as `text` with some new parameters:
- `horizontal_align_down` (`horizontal_align` controls the upper side)
- `space`: controls the space of two sides of text

### `figure`
- keep_aspect_ratio: 0 for not keeping the aspect ratio, 1 for keeping it
- position: [0,0] is the top-left cornor
- width
- height
- horizontal_align
- vertical_align
