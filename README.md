# An Introduction to VBA
*A compact introduction to using VBA.*

Visual Basic for Applications (VBA) is an object based language that is commonly run with a host application, such as Excel, to automate tedious tasks. VBA is related to Visual Basic (VB), with the main difference being that VB can make stand-alone executables, whereas VBA requires a host application. In particular, this guide will focus on using VBA with Excel.

There are plenty of good, comprehensive guides to using VBA available online. The aim of this guide is the opposite - to present VBA in a compact way for someone who is already familiar with at least one programming language.

## Table of Contents



## Introduction

It is entirely possible to make use of VBA without really learning the language, and we will do so in this section by recording macros. However, knowing only this will severely restrict one's ability to create automation for more sophisticated tasks.

### Getting Started

Before doing anything, we should first ensure that we have the correct tools available in the ribbon. For Excel 2007 and newer, add the "Developer" tab. For older than 2007, add the "Control Toolbox" and "Formulas" tabs.

As with an programming language, an IDE is required. For VBA, this is installed by default in the application and can be opened by clicking on "Visual Basic" in the Developer tab.

### Macros

In general, we can create macros to automate tasks in Excel. Macros can be created by either writing a *module* in the IDE, or by recording a macro (which will also be saved as a module).

Indeed, the most basic way to work with VBA is to let Excel create the code by recording a macro. This can be done by using the "Record Macro" option in the Developer tab. Given the straightforward nature of VBA syntax, we can make minor tweaks to the module as we wish.

Generally, VBA code to be run is contained within subroutines,  with naming commonly done in camelCase. Commenting is simply prefaced with `'`.

```VBA
Sub macroName()
' This is a comment.
	Statement1
	Statement2
End Sub
```

To actually run a macro, we can insert a button into the sheet by clicking on "Insert" then "Button (Form Control)" in the Developer tab. We can then assign a macro to the button, and execute the macro by clicking on it. 

Of course, there is only so much that can be achieved without knowing VBA. In the following sections, we will look at creating macros by actually writing VBA code in the IDE. To start writing a module, right-click in the project pane to insert a new module.

Note that we can also use VBA to apply changes to entire sheets or workbooks. In the project pane, simply select the sheet or workbook to work on, instead of inserting a new module. The IDE can then be used to run the code.

## Sheets and Cells

There is a myriad of useful ways to interact with sheets and cells through VBA. In this section, we will look at some commonly used methods, objects and properties. See [here](https://docs.microsoft.com/en-us/office/vba/api/overview/Excel/object-model) for a full list of VBA objects and their corresponding methods and properties.

### Sheets vs Worksheets vs Workbooks

People commonly use "sheets", "worksheets", and even "workbooks" interchangeably. In the context of VBA (or even Excel), it is important to understand the difference between these because these objects all have different methods and properties (although there is some overlap in name and functionality between these objects).

A *sheet* is a collection of *worksheets* and *chart sheets*. A *workbook* is a collection of sheets. In VBA, there is a difference between using the singular and plural form of these objects. For example, if we wanted to use a particular worksheet as a parameter to some function, then we would use the `Worksheet` object rather than the `Worksheets` object.

In particular, VBA contains the [`Sheets` object](https://docs.microsoft.com/en-us/office/vba/api/excel.sheets), the [`Worksheet` object](https://docs.microsoft.com/en-us/office/vba/api/excel.worksheet), the [`Worksheets` object](https://docs.microsoft.com/en-us/office/vba/api/excel.worksheets), the [`Chart` object](https://docs.microsoft.com/en-us/office/vba/api/excel.chart(object)), [`Charts` object](https://docs.microsoft.com/en-us/office/vba/api/excel.charts), the [`Workbook` object](https://docs.microsoft.com/en-us/office/vba/api/excel.workbook), and the [`Workbooks` object](https://docs.microsoft.com/en-us/office/vba/api/excel.workbooks).


### Selections

The `Select` method can be used to select a cell, or multiple cells, in the current sheet.

```VBA
Sub example()
	'Selecting cells H1 and D8.
	Range("H1, D8").Select
End Sub
```

We can also select cells in sheets other than the current one by using the `Activate` method.

```VBA
Sub example()
	'Selecting cells G2, J9 and P10 in Sheet5.
	Sheets("Sheet5").Activate
	Range("G2, J9, P10").Select
End Sub
```

Instead of creating a Range object by selecting cells individually within the `Range` property, we can select a range of them, or even a range that has been renamed.

```VBA
Sub example()
	'Selecting cells A1 to A5.
	Range("A1:A5").Select
	
	'Selecting cells from the namedRange range.
	Range("namedRange").Select
End Sub
```

Similarly, we can select entire rows or columns by using the `Rows` or `Columns` properties, respectively. Alternatively, we can still use the `Range` property.

```VBA
Sub example()
	'Selecting rows 2 to 7.
	Range("2:7").Select
	Rows("2:7").Select
	
	`Selecting columns B to H.
	Range("B:H").Select
	Columns("B:H").Select
End Sub
```

Quite often we will want to be able to dynamically select cells, rather than a particular one. For example, we may want to select the *n*th row in the *m*th column, where *n* and *m* depend on what's happened in the module. To do this, we can use the `Cells` property.

```VBA
Sub example()
	'Selecting the cell in row 4, column 7.
	Cells(4, 7).Select
End Sub
```

The `Offset` property can be used to move a selection.

```VBA
Sub example()
	'Moving the selected cell down 2, right 7.
	ActiveCell.Offset(2, 7).Select
End Sub
```

### Properties

VBA uses *dot notation* to denote hierarchy when manipulating an object. For example, if we wanted to manipulate the size of the font of a Range object, then we could use `ActiveCell.Font.Size`.

####  Cell Content Manipulation

The `Value` property represents the contents of a cell. This can be used to change the contents of a cell.

```VBA
Sub example()
	'Setting the value of cell B3 to 19.
	Range("B3").Value = 19
	
	'Setting the value of cell D9 to the string "This is some text".
	Range("D9").Value = "This is some text"
End Sub
```

Of course, we can use dot notation to change the contents of cells in more specific locations.

```VBA
Sub example()
	'Setting the value of cell B3 in Sheet2 of another open workbook to 19.
	Workbook("Book5.xlsx").Sheets("Sheet2").Range("B3").Value = 19
End Sub
```

Note that not using the `Value` property would have the same effect, since if no other property is specified, then the value of the cell is modified by default.

```VBA
Sub example()
	'Setting the value of cell B3 to 19.
	Range("B3") = 19
End Sub
```

Of course, it is also possible to change the value (or any other property, such as font size) of a cell based on another cell, or even itself.

```VBA
Sub example()
	'Setting the value of cell B3 to the value of cell A1.
	Range("B3") = Range("Al")

	'Increasing the value of cell D2 by 1, every time the macro is run.
	Range("D2") = Range("D2") + 1
End Sub
```

The `ClearContents` method can be used to erase the contents of a cell.

```VBA
Sub example()
	'Erasing the contents of cell B3.
	Range("B3").ClearContents
End Sub
```

#### Text Formatting

To format text we will access the `Font` property. Within the IDE simply typing `Range("A1").Font.` will reveal a list of properties belonging to the `Font` property. For reference, we will list some of these below.

The `Size` property can be used to change the text size.

```VBA
Sub example()
	'Formatting the contents of cell B3 to have font size 18.
	Range("B3").Font.Size = 18
End Sub
```

The `Bold`, `Italic`, and `Underline` properties can be used to give text the bold, italic, and underline emphasis, respectively.

```VBA
Sub example()
	'Formatting the contents of cell B3 to be in bold.
	Range("B3").Font.Bold = True
	
	'Formatting the contents of cell D6 to be in italics.
	Range("D6").Font.Italics = True
	
	'Formatting the contents of cell A1 to be underlined.
	Range("A1").Font.Underline = True
End Sub
```

The `Name` property can be used to set the font style.

```VBA
Sub example()
	'Formatting the contents of cell B3 to have the Arial font.
	Range("B3").Font.Name = "Arial"
End Sub
```

#### Borders
Here are a couple of other commonly used properties.

The `Borders` property can be used to add a border to cells. Similar to the `Font` property, we can use the IDE to reveal a list of properties belonging to the `Borders` property.

```VBA
Sub example()
	'Adding a border to cells B3 to B9.
	Range("B3:B9").Borders.Value = 1

	'Making the borders as thick as possible.
	Range("B3:B9").Borders.Weight = 4

	'Removing borders from cell A2.
	Range("A2").Borders.Value = 0
End Sub
```

#### The `With` Statement

Suppose we wanted to change various properties of some cells.

```VBA
Sub example()
	Range("B3:B9").Borders.Weight = 3
	Range("B3:B9").Font.Bold = True
	Range("B3:B9").Font.Size = 18
	Range("B3:B9").Font.Italic = True
	Range("B3:B9").Font.Name = "Arial"
End Sub
```

We can reduce repeated code by using the `With` statement.

```VBA
Sub example()
	With Range("B3:B9")
		.Borders.Weight = 3
		.Font.Bold = True
		.Font.Size = 18
		.Font.Italic = True
		.Font.Name = "Arial"
	End With
End Sub
```

We can even go further and reduce on the repetition of `.Font`.

```VBA
Sub example()
	With Range("B3:B9")
		.Borders.Weight = 3
		With .Font
			.Bold = True
			.Size = 18
			.Italic = True
			.Name = "Arial"
		End With
	End With
End Sub
```

### Colors

In the previous section we looked at the `Font` and `Border` properties. These properties can also be seen as objects with `Color`, or `ColorIndex`, as a property. In this section we will look at manipulating the color of cells and worksheet tabs.

Colors can be set either by using the `ColorIndex` property, which is preferred on versions of Excel older than 2007, or the `Color` property, which provides the full range of colors. Both of these are properties of the Font, Border, and Interior objects, and also the `Tab` property.

#### `ColorIndex`

Unfortunately, the `ColorIndex` property is limited to only 56 colors, and depends on the color theme of the application. The color-index values for the default color theme can be seen [here](https://docs.microsoft.com/en-us/office/vba/api/excel.colorindex).

```VBA
Sub example()
	'Setting the color of the B3 cell to blue, and the font to white.
	Range("B3").Interior.ColorIndex = 5
	Range("B3").Font.ColorIndex = 2

	'Adding a green border to the B3 to B9 cells.
	Range("B3:B9").Border.ColorIndex = 4

	'Setting the color of the Sheet7 tab to orange.
	Sheets("Sheet7").Tab.ColorIndex = 45
End Sub
```

#### `Color`

Using the `Color` property is similar to using the `ColorIndex` property, except it uses RGB color codes. Attempting to use this on versions of Excel older than 2007 will result in an approximate color being chosen from the color palette of 56 colors.

```VBA
Sub example()
	'Setting the color of the B3 cell to blue, and the font to white.
	Range("B3").Interior.Color = RGB(0, 0, 255)
	Range("B3").Font.Color = RGB(255, 255, 255)

	'Adding a green border to the B3 to B9 cells.
	Range("B3:B9").Border.Color = RGB(0, 255, 0)

	'Setting the color of the Sheet7 tab to orange.
	Sheets("Sheet7").Tab.Color = RGB(255, 128, 0)
End Sub
```

## Variables

By default VBA behaves like a statically-typed language, which means that variables need to be given a type before usage. It is possible to use the [`Infer` statement](https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/option-infer-statement) to change VBA's behaviour to a dynamically-typed one, but this comes with the usual disadvantage of more memory usage and slower runtime.

## Conditions

## Loops

## Procedures and Functions

## Dialog Boxes

## Events

## Forms and Controls

## Arrays
