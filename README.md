# Unpivot Columns | An Excel Macro

An excel macro to unpivot columns.

### Table of Contents

- [Unpivotting a Column](#unpivotting-a-column)
- [Installation](#installation)
- [Tutorial](#tutorial)
- [Contributing](#contributing)
- [License and Credits](#license-and-credits)

### Unpivotting a Column

![](https://github.com/Levantino-Engineering/unpivot-columns-excel/blob/main/screenshots/unpivot_column_definition.png)

Source: https://support.microsoft.com/en-us/office/unpivot-columns-power-query-0f7bad4b-9ea1-49c1-9d95-f588221c7098

### Installation

0. Open or create the macro-enabled workbook that you want to import the macro. 
1. Enable the Developer tab on the File tab. Then Options/Customize Ribbon/Developer.
2. Open the VBE, either by pressing Alt + F11 on your keyboard or by opening the Visual Basic command in the Developer tab.
3. Right-click the VBAProject with your workbook's name, then select the option Import File.
4. Select the .bas file.

The macro has been successfully installed!

### Tutorial

In order for the macro to work, follow these instructions:

0. Move the table you want to unpivot to an empty sheet
1. Make sure the attributes are at the first row and start at the A2 cell. The cell A1 must be empty.

Here is an example:

![](https://github.com/Levantino-Engineering/unpivot-columns-excel/blob/main/screenshots/example/before_unpivot.png)

There are two ways to run the macro:

- By using the default assigned shorcut CTRL + SHIFT + P. You can change the shorcut by selecting the Macro command located at the Developer tab, then pressing Options.
- By running the macro on the Macro command located at the Developer tab.

Here is the result:

![](https://github.com/Levantino-Engineering/unpivot-columns-excel/blob/main/screenshots/example/after_unpivot.png)

### Contributing

Contributions are welcome! Please review the contribution guidelines on how to:

- Report issues
- File pull requests
- Support the project as a non-developer

### License and Credits

*Copyright © 2023 [Alejandro Sánchez](https://github.com/Levantino-Engineering) (Levantino Engineering)*

Licensed under the _GNU AGPLv3_, extended by a number of additional terms. This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY. For more information on the license please see the [LICENSE file](https://github.com/Levantino-Engineering/unpivot-columns-excel/blob/main/LICENSE.txt) accompanying this add-on. The source code is available on [GitHub](https://github.com/Levantino-Engineering).

----------------------------------------------------------

Return to the [Table Of Contents](#table-of-contents)
