# An Introduction to VBA
*A compact introduction to using VBA.*

Visual Basic for Applications (VBA) is an object-based language (but isn't an OOP language, since it doesn't support inheritance) that is commonly run with a host application, such as Excel, to automate tedious tasks. VBA is related to Visual Basic (VB), with the main difference being that VB can make stand-alone executables, whereas VBA requires a host application. In particular, this guide will focus on using VBA with Excel.

There are plenty of good, comprehensive guides to using VBA available online. The aim of this guide is the opposite - to present VBA in a compact way for someone who is already familiar with at least one programming language.

## Table of Contents



## Introduction

It is entirely possible to make use of VBA without really learning the language, and we will do so in this section by recording macros. However, knowing only this will severely restrict one's ability to create automation for more sophisticated tasks.

### Getting Started

Before doing anything, we should first ensure that we have the correct tools available in the ribbon. For Excel 2007 and newer, add the "Developer" tab. For older than 2007, add the "Control Toolbox" and "Formulas" tabs.

As with any programming language, an IDE is required. For VBA, this is installed by default in the application and can be opened by clicking on "Visual Basic" in the Developer tab.

### Macros

In general, we can create macros to automate tasks in Excel. Macros can be created by either writing explicit VBA code in the IDE, or by recording a macro.

Indeed, the most basic way to work with VBA is to let Excel create the code by recording a macro. This can be done by using the "Record Macro" option in the Developer tab. Given the straightforward nature of VBA syntax, we can make minor tweaks to the code as we wish.

Generally, VBA code to be run is contained within *subroutines* (also called *procedures*),  with naming commonly done in camelCase. Commenting is simply prefaced with a `'` followed by a space.

```VBA
Sub macroName()
	' This is a comment.
	Statement1
	Statement2
End Sub
```

To actually run a macro, we can either use the IDE itself to run the code, or we can insert a button into the sheet by clicking on "Insert" then "Button (Form Control)" in the Developer tab. We can then assign a macro to the button, and execute the macro by clicking on it.

Of course, there is only so much that can be achieved without knowing VBA. In the following sections, we will look at creating macros by actually writing VBA code in the IDE.

### Modules

A *module* is a code container. The main types of modules are:
* Standard Code Modules: also called Code Modules, or just Modules. This is where most VBA code should go unless there is good reason to use another module type.
* Workbook and Sheet Modules: contain VBA code which control event procedures for workbooks and sheets.
* UserForm Modules: contain VBA code which controls UserForm objects.
* Class Modules: contain VBA code used to create new VBA objects.

The IDE project pane can be used to write code in the workbook and sheet modules, which are created automatically, or to insert the other module types.

It is important to correctly choose which module type to use in order to avoid unexpected results and to maintain high levels of code hygiene. For example, using sheet modules can create unexpected results when the sheet itself is deleted, copied or moved. On the other hand, using standard code modules allows for the logical structuring of code as units, which can then be version controlled and managed easily in large project. Of course, not all code should blindly be put into standard code modules - event procedures put into a standard code module will fail to execute.

Throughout the rest of this guide, unless otherwise specified, we will assume that all VBA code is placed in a standard code module. Hence, all usage of the word "module" will refer to standard code modules. The other three module types are not the main focus of this guide, but workbook and sheet modules are used in the [Events](#events) section, UserForm modules are used in the [Forms and Controls](#forms-and-controls) section, while class modules are not used at all.

## Sheets and Cells

There is a myriad of useful ways to interact with sheets and cells through VBA. In this section, we will look at some commonly used methods, objects and properties. See [here](https://docs.microsoft.com/en-us/office/vba/api/overview/Excel/object-model) for a full list of VBA objects and their corresponding methods and properties.

### Sheets vs Worksheets vs Workbooks

People commonly use "sheets", "worksheets", and even "workbooks" interchangeably. In the context of VBA (or even Excel), it is important to understand the difference between these because these objects all have different methods and properties (although there is some overlap in name and functionality between these objects).

A *sheet* is a collection of *worksheets* and *chart sheets*. A *workbook* is a collection of sheets. In VBA, there is a difference between using the singular and plural form of these objects. For example, if we wanted to use a particular worksheet as a parameter to some function, then we would use the `Worksheet` object rather than the `Worksheets` object.

In particular, VBA contains the [`Sheets` object](https://docs.microsoft.com/en-us/office/vba/api/excel.sheets), the [`Worksheet` object](https://docs.microsoft.com/en-us/office/vba/api/excel.worksheet), the [`Worksheets` object](https://docs.microsoft.com/en-us/office/vba/api/excel.worksheets), the [`Chart` object](https://docs.microsoft.com/en-us/office/vba/api/excel.chart(object)), [`Charts` object](https://docs.microsoft.com/en-us/office/vba/api/excel.charts), the [`Workbook` object](https://docs.microsoft.com/en-us/office/vba/api/excel.workbook), and the [`Workbooks` object](https://docs.microsoft.com/en-us/office/vba/api/excel.workbooks).


### Selections

The `Select` method can be used to select a cell, or multiple cells, in the current sheet.

```VBA
' Selecting cells H1 and D8.
Range("H1, D8").Select
```

We can also select cells in sheets other than the current one by using the `Activate` method.

```VBA
' Selecting cells G2, J9 and P10 in Sheet5.
Sheets("Sheet5").Activate
Range("G2, J9, P10").Select
```

Instead of creating a Range object by selecting cells individually within the `Range` property, we can select a range of them, or even a range that has been renamed.

```VBA
' Selecting cells A1 to A5.
Range("A1:A5").Select

' Selecting cells from the namedRange range.
Range("namedRange").Select
```

Similarly, we can select entire rows or columns by using the `Rows` or `Columns` properties, respectively. Alternatively, we can still use the `Range` property.

```VBA
' Selecting rows 2 to 7.
Range("2:7").Select
Rows("2:7").Select

' Selecting columns B to H.
Range("B:H").Select
Columns("B:H").Select
```

Quite often we will want to be able to dynamically select cells, rather than a particular one. For example, we may want to select the *n*th row in the *m*th column, where *n* and *m* depend on what's happened in the module. To do this, we can use the `Cells` property.

```VBA
' Selecting the cell in row 4, column 7.
Cells(4, 7).Select
```

The `Offset` property can be used to move a selection.

```VBA
' Moving the selected cell down 2, right 7.
ActiveCell.Offset(2, 7).Select
```

### Properties

VBA uses *dot notation* to denote hierarchy when manipulating an object. For example, if we wanted to manipulate the size of the font of a Range object, then we could use `ActiveCell.Font.Size`.

####  Cell Content Manipulation

The `Value` property represents the contents of a cell. This can be used to change the contents of a cell.

```VBA
' Setting the value of cell B3 to 19.
Range("B3").Value = 19

' Setting the value of cell D9 to the string "This is some text".
Range("D9").Value = "This is some text"
```

Of course, we can use dot notation to change the contents of cells in more specific locations.

```VBA
' Setting the value of cell B3 in Sheet2 of another open workbook to 19.
Workbook("Book5.xlsx").Sheets("Sheet2").Range("B3").Value = 19
```

Note that not using the `Value` property would have the same effect, since if no other property is specified, then the value of the cell is modified by default.

```VBA
' Setting the value of cell B3 to 19.
Range("B3") = 19
```

Of course, it is also possible to change the value (or any other property, such as font size) of a cell based on another cell, or even itself.

```VBA
' Setting the value of cell B3 to the value of cell A1.
Range("B3") = Range("Al")

' Increasing the value of cell D2 by 1, every time the macro is run.
Range("D2") = Range("D2") + 1
```

The `ClearContents` method can be used to erase the contents of a cell.

```VBA
' Erasing the contents of cell B3.
Range("B3").ClearContents
```

#### Text Formatting

To format text we will access the `Font` property. Within the IDE simply typing `Range("A1").Font.` will reveal a list of properties belonging to the `Font` property. For reference, we will list some of these below.

The `Size` property can be used to change the text size.

```VBA
' Formatting the contents of cell B3 to have font size 18.
Range("B3").Font.Size = 18
```

The `Bold`, `Italic`, and `Underline` properties can be used to give text the bold, italic, and underline emphasis, respectively.

```VBA
' Formatting the contents of cell B3 to be in bold.
Range("B3").Font.Bold = True

' Formatting the contents of cell D6 to be in italics.
Range("D6").Font.Italics = True

' Formatting the contents of cell A1 to be underlined.
Range("A1").Font.Underline = True
```

The `Name` property can be used to set the font style.

```VBA
' Formatting the contents of cell B3 to have the Arial font.
Range("B3").Font.Name = "Arial"
```

#### Borders
Here are a couple of other commonly used properties.

The `Borders` property can be used to add a border to cells. Similar to the `Font` property, we can use the IDE to reveal a list of properties belonging to the `Borders` property.

```VBA
' Adding a border to cells B3 to B9.
Range("B3:B9").Borders.Value = 1

' Making the borders as thick as possible.
Range("B3:B9").Borders.Weight = 4

' Removing borders from cell A2.
Range("A2").Borders.Value = 0
```

#### The `With` Statement

Suppose we wanted to change various properties of some cells.

```VBA
Range("B3:B9").Borders.Weight = 3
Range("B3:B9").Font.Bold = True
Range("B3:B9").Font.Size = 18
Range("B3:B9").Font.Italic = True
Range("B3:B9").Font.Name = "Arial"
```

We can reduce repeated code by using the `With` statement.

```VBA
With Range("B3:B9")
	.Borders.Weight = 3
	.Font.Bold = True
	.Font.Size = 18
	.Font.Italic = True
	.Font.Name = "Arial"
End With
```

We can even go further and reduce on the repetition of `.Font`.

```VBA
With Range("B3:B9")
	.Borders.Weight = 3
	With .Font
		.Bold = True
		.Size = 18
		.Italic = True
		.Name = "Arial"
	End With
End With
```

### Colours

In the previous section we looked at the `Font` and `Border` properties. These properties can also be seen as objects with `Color`, or `ColorIndex`, as a property. In this section we will look at manipulating the colour of cells and worksheet tabs.

Colours can be set either by using the `ColorIndex` property, which is preferred on versions of Excel older than 2007, or the `Color` property, which provides the full range of colours. Both of these are properties of the Font, Border, and Interior objects, and also the `Tab` property.

#### `ColorIndex`

Unfortunately, the `ColorIndex` property is limited to only 56 colours, and depends on the colour theme of the application. The colour-index values for the default colour theme can be seen [here](https://docs.microsoft.com/en-us/office/vba/api/excel.colorindex).

```VBA
' Setting the colour of the B3 cell to blue, and the font to white.
Range("B3").Interior.ColorIndex = 5
Range("B3").Font.ColorIndex = 2

' Adding a green border to the B3 to B9 cells.
Range("B3:B9").Border.ColorIndex = 4

' Setting the colour of the Sheet7 tab to orange.
Sheets("Sheet7").Tab.ColorIndex = 45
```

#### `Color`

Using the `Color` property is similar to using the `ColorIndex` property, except it uses RGB colour codes. Attempting to use this on versions of Excel older than 2007 will result in an approximate colour being chosen from the colour palette of 56 colours.

```VBA
' Setting the colour of the B3 cell to blue, and the font to white.
Range("B3").Interior.Color = RGB(0, 0, 255)
Range("B3").Font.Color = RGB(255, 255, 255)

' Adding a green border to the B3 to B9 cells.
Range("B3:B9").Border.Color = RGB(0, 255, 0)

' Setting the colour of the Sheet7 tab to orange.
Sheets("Sheet7").Tab.Color = RGB(255, 128, 0)
```

## Variables and Types

By default, VBA behaves like a statically-typed language, which means that variables need to be given a type before usage. It is possible to use the [`Infer` statement](https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/option-infer-statement) to change VBA's behaviour to a dynamically-typed one, but this comes with the usual disadvantage of more memory usage and slower runtime.

### Types

See the [documentation on types](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/data-type-summary). Commonly used types for typical programming are the Boolean, Double, Integer and String types.

### Variables

Variable declaration is done as usual (often named in camelCase).

```VBA
Dim variableName As Type
```

We may also declare multiple variables in a single line.

```VBA
Dim variableName1 As Type1, variableName2 As Type2, variableName3 As Type3
```

### Variable Scope

As usual, the scope of a variable is, by default, the subroutine in which it is declared.

We can declare variables at the beginning of a module to allow usage within the entire module.

```VBA
Dim variableName As Type

Sub numberOne()
	' We can use variableName here.
End Sub

Sub numberTwo()
	' We can also use variableName here.
End Sub
```

If we want to use variables globally across all modules, then we can instead use the `Global` statement at the beginning of a module.

```VBA
Global variableName As Type
```

The `Static` statement can be used to allow the value of a variable to persist.

```VBA
Sub numberOne()
	Static variableName As Type
End Sub
```

We can also allow the value of all variables within a subroutine to persist.

```VBA
Static Sub numberOne()
	Dim variableName1 As Type1, variableName2 As Type2, variableName3 As Type3
End Sub
```

### Constants

Constants can’t be changed later and are declared similarly to variables, but can only be Boolean, Byte, Integer, Long, Currency, Single, Double, Date, String, or Variant types. Since the value of a constant should already be known, we can specify the value and type in a single line.

```VBA
Const gravity As Double = 9.80665
```

## Conditionals

VBA has the usual comparison and logical operators: `=`, `<>`, `<`, `>`, `<=`, `>=`, `And`, `Or`, `Not`, `Xor`.

### The `If`, `ElseIf` and `Else` Statements

As usual, a block of code is executed given that a condition is true.

```VBA
If condition1 Then
	statement1
ElseIf condition2 Then
	statement2
Else
	statement3
End If
```

### The `Select Case` Statement

We can use a `Select Case` statement to handle cases more elegantly.

```VBA
Dim x As Integer
Select Case x
Case Is = 5
	xComment = "This is the maximum achievable!"
Case Is = 4
	xComment = "This is almost the maximum achievable!"
Case Else
	xComment = "This is quite average"
End Select
```

### Wildcard Characters and the `Like` Operator

When working with a lot of data, it is common to want to be able to compare whether the contents of one cell is contained in the contents of another cell. Since the comparison operators are not capable of doing this, VBA uses wildcard characters with the `Like` operator.

The wildcard usage is typical and can be read in detail [here](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/wildcard-characters-used-in-string-comparisons).

```VBA
Dim exampleString As String
exampleString = "Example 12345"

' All of the following conditionals return True:

exampleString Like "*12345*"
exampleString Like "Example 12###"
exampleString Like "?xample?1234?"
exampleString Like "[DEF]xample 1234[4-7]"
exampleString Like "[!GHIJ]xample 1234[!6-9]"
```

## Loops
VBA has three looping constructs - the `While`, `Do`, and `For` loops.

### The `While` Loop

The basic syntax for a `While` loop is as follows:

```VBA
While condition
	statement
Wend
```

### The `Do` Loops

VBA has four different syntaxes for do loops. These loops will repeat a block of statements while a condition is `True`, or until a condition becomes `True`. By placing the `While` or `Until` conditions in different places, we can produce different outcomes. For example, by placing the `While` or `Until` conditions at the end, we can guarantee that any statements will execute at least once.

```VBA
Do While condition
	statement
Loop
```

```VBA
Do
	statement
Loop While condition
```

```VBA
Do Until condition
	statement
Loop
```

```VBA
Do
	statement
Loop Until condition
```

### The `For` Loop

The basic syntax for a `For` loop is as follows:

```VBA
For i = startNumber To endNumber
	statement
Next
```

## Subroutines and Functions

We have already seen in the Introduction section that VBA code to be run is contained within subroutines in the following way:

```VBA
Sub macroName()
	statement
End Sub
```

In this section we will look at different types of subroutines; some different functionalities of subroutines; and the difference between a subroutine and a function.

### Public and Private Subroutines

A Public subroutine is one which can be accessed from any module. This can be specified by using the `Public` statement, however, not including the statement will produce the same effect. Hence, the following two syntaxes are equivalent:

```VBA
Public Sub example()
	statement
End Sub
```

```VBA
Sub example()
	statement
End Sub
```

A Private subroutine is one which can only be accessed from within the module in which it was created. This is useful to keep namespaces tidy.

```VBA
Private Sub example()
	statement
End Sub
```

### Calling Subroutines

Much like normal functions in a typical programming language, subroutines can be called within other subroutines.

```VBA
Sub subroutine1()
	statement
End Sub

Sub subroutine2()
	subroutine1
End Sub
```

### Arguments

Also like normal functions in a typical programming language, subroutines can also accept arguments, with the usual [variable scope](#variable-scope) rules applying.

```VBA
Sub subroutine1(variableName1 As Type1, variableName2 As Type2, variableName3 As Type3)
	statement ' This can be dependent on variableName1, variableName2 and variableName3.
End Sub

Sub subroutine2()
	subroutine1 someArgument1, someArgument2, someArgument3
End Sub
```

By default, if a subroutine accepts arguments, then they are not optional - a subroutine called without the correct number of arguments will fail to execute. The `Optional` keyword can be used to specify optional arguments, but these must be specified after any non-optional arguments.

```VBA
Sub subroutine1(variableName1 As Type1, Optional variableName2 As Type2, Optional variableName3 As Type3)
	statement ' This can be dependent on variableName1, variableName2 and variableName3.
End Sub

Sub subroutine2()
	' All of the following subroutine calls are valid:
	subroutine1 someArgument1
	subroutine1 someArgument1, someArgument2
	subroutine1 someArgument1, someArgument2, someArgument3
End Sub
```

### Passing Arguments by Value and by Reference

By default,  VBA is a pass-by-reference language. We now go into detail about what this means.

When passing arguments to a subroutine, we can choose to pass them either by value or by reference:
* Passing by reference: done by using the `ByRef` keyword. In VBA this is the default way to pass an argument if nothing is specified. In particular, this means that if a variable is passed as an argument, then its reference is transmitted.
* Passing by value: done by using the `ByVal` keyword. In particular, this means that if a variable is passed as an argument, then only its value is transmitted.

The differences are best demonstrated by example.

```VBA
' Creating two subroutines which squares numbers, with different pass bys.
Sub calculateSquare1(ByRef number As Integer)
	number = number * number
End Sub

Sub calculateSquare2(ByVal number As Integer)
	number = number ^ 2
End Sub

' Testing out the two subroutines.
Sub test()
	Dim testNumber1 As Integer, testNumber2 As Integer
	testNumber1 = 20
	testNumber2 = 20
	
	calculateSquare1 testNumber1
	calculateSquare2 testNumber2

	' The value of testNumber1 will be displayed as 400.
	MsgBox testNumber1
	
	' The value of testNumber2 will be displayed as 20.
	MsgBox testNumber2
End Sub
```

Deciding between passing by value or passing by reference depends on both variable protection and performance. If we want to prevent a variable from being modified, then we should use `ByVal`. However, since passing by value works by copying the entire data contents of the variable, a large value type would be more efficiently passed by reference. See [here](https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/language-features/procedures/passing-arguments-by-value-and-by-reference) for more details.

### Functions

In VBA, subroutines and functions are similar, except functions must return a value. Using the function name within the function itself is used to determine the return value. Subroutines and functions can be called inside of each other, but function calls should have arguments enclosed in `()`.

```VBA
' Creating a function which squares numbers.
Function calculateSquare(number As Integer)
	calculateSquare = number ^ 2
End Function

' Testing out the calculateSquare function.
Sub test()
	result = calculateSquare(20)

	' The value of result will be displayed as 400.
	MsgBox result
End Sub
```

Functions which are created inside modules can be used in a worksheet just like any other Excel function.

## Dialog Boxes

## Events

## Forms and Controls

## Arrays
