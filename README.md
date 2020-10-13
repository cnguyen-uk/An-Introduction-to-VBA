# An Introduction to VBA
*A compact introduction to using VBA.*

Visual Basic for Applications (VBA) is an object-based language (but isn't an OOP language, since it doesn't support inheritance) that is commonly run with a host application, such as Excel, to automate tedious tasks. VBA is related to Visual Basic (VB), with the main difference being that VB can make stand-alone executables, whereas VBA requires a host application. In particular, this guide will focus on using VBA with Excel.

There are plenty of good, comprehensive guides to using VBA available online. The aim of this guide is the opposite - to present VBA in a compact way for someone who is already familiar with at least one programming language.

## Table of Contents

- [Introduction](#introduction)
  * [Getting Started](#getting-started)
  * [Macros](#macros)
  * [Lowering Macro Execution Time](#lowering-macro-execution-time)
  * [Modules](#modules)
  * [Printing](#printing)
- [Sheets and Cells](#sheets-and-cells)
  * [Sheets vs Worksheets vs Workbooks](#sheets-vs-worksheets-vs-workbooks)
  * [Selections](#selections)
  * [Properties](#properties)
    + [Cell Content Manipulation](#cell-content-manipulation)
    + [Text Formatting](#text-formatting)
    + [Borders](#borders)
    + [The `With` Statement](#the-with-statement)
  * [Colours](#colours)
    + [The `ColorIndex` Property](#the-colorindex-property)
    + [The `Color` Property](#the-color-property)
- [Variables and Types](#variables-and-types)
  * [Types](#types)
  * [Variables](#variables)
  * [Variable Scope](#variable-scope)
  * [Constants](#constants)
- [Conditionals](#conditionals)
  * [The `If`, `ElseIf` and `Else` Statements](#the-if-elseif-and-else-statements)
  * [The `Select Case` Statement](#the-select-case-statement)
  * [Wildcard Characters and the `Like` Operator](#wildcard-characters-and-the-like-operator)
- [Loops](#loops)
  * [The `While` Loop](#the-while-loop)
  * [The `Do` Loops](#the-do-loops)
  * [The `For` Loop](#the-for-loop)
  * [The `For Each` Loop](#the-for-each-loop)
- [Subroutines and Functions](#subroutines-and-functions)
  * [Public and Private Subroutines](#public-and-private-subroutines)
  * [Calling Subroutines](#calling-subroutines)
  * [Arguments](#arguments)
  * [Passing Arguments by Value and by Reference](#passing-arguments-by-value-and-by-reference)
  * [Functions](#functions)
- [Dialog Box Functions](#dialog-box-functions)
  * [The `MsgBox()` Function](#the-msgbox-function)
  * [The `InputBox()` Function](#the-inputbox-function)
- [Events](#events)
  * [Workbook Events](#workbook-events)
  * [Worksheet Events](#worksheet-events)
  * [Deactivating Events](#deactivating-events)
- [Forms and Controls](#forms-and-controls)
  * [UserForms](#userforms)
    + [Events](#events-1)
    + [Launching](#launching)
  * [Controls](#controls)
    + [Label](#label)
    + [TextBox](#textbox)
    + [CommandButton](#commandbutton)
    + [CheckBox](#checkbox)
    + [OptionButton](#optionbutton)

## Introduction

It is entirely possible to make use of VBA without really learning the language, and we will do so in this section by recording macros. However, knowing only this will severely restrict one's ability to create automation for more sophisticated tasks.

### Getting Started

Before doing anything, we should first ensure that we have the correct tools available in the ribbon. For Excel 2007 and newer, add the "Developer" tab. For older than 2007, add the "Control Toolbox" and "Formulas" tabs.

As with any programming language, an IDE is required. For VBA, this is installed by default in the application and can be opened by clicking on "Visual Basic" in the Developer tab.

### Macros

In general, we can create macros to automate tasks in Excel. Macros can be created by either writing explicit VBA code in the IDE, or by recording a macro.

Indeed, the most basic way to work with VBA is to let Excel create the code by recording a macro. This can be done by using the "Record Macro" option in the Developer tab. Given the straightforward nature of VBA syntax, we can make minor tweaks to the code as we wish.

Generally, VBA code to be run is contained within *subroutines* (also called *procedures*),  with naming commonly done in camelCase. Commenting is simply prefaced with a `'`.

```VBA
Sub macroName()
    ' This is a comment.
    statement1
    statement2
End Sub
```

To actually run a macro, we can either use the IDE itself to run the code, or we can insert a button into the sheet by clicking on "Insert" then "Button (Form Control)" in the Developer tab. We can then assign a macro to the button, and execute the macro by clicking on it.

Of course, there is only so much that can be achieved without knowing VBA. In the upcoming sections, we will look at creating macros by actually writing VBA code in the IDE.

### Lowering Macro Execution Time

If a macro results in a lot of modifications to a workbook, then Excel will update the workbook display for every modification. This can severely reduce the speed of the macro. The following code will tell Excel to not update the display, and hence increase the macro's speed:

```VBA
Application.ScreenUpdating = False
statement
Application.ScreenUpdating = True
```

### Modules

A *module* is a code container. The main types of modules are:
* Standard Code Modules: also called Code Modules, or just Modules. This is where most VBA code should go unless there is good reason to use another module type.
* Workbook and Sheet Modules: contain VBA code which control event subroutines for workbooks and sheets.
* UserForm Modules: contain VBA code which controls UserForm objects.
* Class Modules: contain VBA code used to create new VBA objects.

The IDE Project Explorer can be used to choose which module to write code into. The workbook and sheet modules are created automatically, but we can also insert the other module types.

It is important to correctly choose which module type to use in order to avoid unexpected results and to maintain high levels of code hygiene. For example, using sheet modules can create unexpected results when the sheet itself is deleted, copied or moved. On the other hand, using standard code modules allows for the logical structuring of code as units, which can then be version controlled and managed easily in large project. Of course, not all code should blindly be put into standard code modules - event subroutines put into a standard code module will fail to execute.

Throughout the rest of this guide, unless otherwise specified, we will assume that all VBA code is placed in a standard code module. Hence, all usage of the word "module" will refer to standard code modules. The other three module types are not the main focus of this guide, but workbook and sheet modules are used in the [Events](#events) section, UserForm modules are used in the [Forms and Controls](#forms-and-controls) section, while class modules are not used at all.

### Printing

As with any programming language, it is useful to be able to print outputs. In VBA there are several ways to do this, such as by outputting to a cell in a worksheet; displaying the output using a dialog box; or by using the Immediate window.

The first two ways are discussed later on in this guide, [here](#cell-content-manipulation) and [here](#the-msgbox-function), and require the least set up to do. Some programmers will prefer to use the third way since it offers a more familiar coding experience. To set up the Immediate window next to the IDE, see [here](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/immediate-window).

To print to the Immediate window, we can use the `Print()` method:

```VBA
Sub helloWorld()
    Debug.Print("Hello World!")
End Sub
```


## Sheets and Cells

There is a myriad of useful ways to interact with sheets and cells through VBA. In this section, we will look at some commonly used methods, objects and properties. See [here](https://docs.microsoft.com/en-us/office/vba/api/overview/Excel/object-model) for a full list of VBA objects and their corresponding methods and properties.

### Sheets vs Worksheets vs Workbooks

People commonly use "sheets", "worksheets", and even "workbooks" interchangeably. In the context of VBA (or even Excel), it is important to understand the difference between these because these objects all have different methods and properties (although there is some overlap in name and functionality between these objects).

A *sheet* is a collection of *worksheets* and *chart sheets*. A *workbook* is a collection of sheets. In VBA, there is a difference between using the singular and plural form of these objects. For example, if we wanted to use a particular worksheet as a parameter to some function, then we would use the `Worksheet` object rather than the `Worksheets` object.

In particular, VBA contains the [`Sheets` object](https://docs.microsoft.com/en-us/office/vba/api/excel.sheets), the [`Worksheet` object](https://docs.microsoft.com/en-us/office/vba/api/excel.worksheet), the [`Worksheets` object](https://docs.microsoft.com/en-us/office/vba/api/excel.worksheets), the [`Chart` object](https://docs.microsoft.com/en-us/office/vba/api/excel.chart(object)), [`Charts` object](https://docs.microsoft.com/en-us/office/vba/api/excel.charts), the [`Workbook` object](https://docs.microsoft.com/en-us/office/vba/api/excel.workbook), and the [`Workbooks` object](https://docs.microsoft.com/en-us/office/vba/api/excel.workbooks).

### Selections

The `Select` method can be used to select a cell, or multiple cells, in the current sheet.

```VBA
Range("H1, D8").Select  ' Select cells H1 and D8
```

We can also select cells in sheets other than the current one by using the `Activate` method.

```VBA
' This statement selects cells G2, J9 and P10 in Sheet5.
Sheets("Sheet5").Activate
Range("G2, J9, P10").Select
```

Instead of creating a Range object by selecting cells individually within the `Range` property, we can select a range of them, or even a range that has been renamed.

```VBA
Range("A1:A5").Select  ' Select cells A1 to A5

Range("namedRange").Select  ' Select cells from the namedRange range
```

Similarly, we can select entire rows or columns by using the `Rows` or `Columns` properties, respectively. Alternatively, we can still use the `Range` property.

```VBA
' Both of the following statements will select rows 2 to 7:
Range("2:7").Select
Rows("2:7").Select

' Both of the following statements will select columns B to H:
Range("B:H").Select
Columns("B:H").Select
```

Quite often we will want to be able to dynamically select cells, rather than a particular one. For example, we may want to select the *n*th row in the *m*th column, where *n* and *m* depend on what's happened in the module. To do this, we can use the `Cells` property.

```VBA
Cells(4, 7).Select  ' Select the cell in row 4, column 7
```

The `Offset` property can be used to move a selection.

```VBA
ActiveCell.Offset(2, 7).Select  ' Move the selected cell down 2, right 7
```

### Properties

VBA uses *dot notation* to denote hierarchy when manipulating an object. For example, if we wanted to manipulate the size of the font of a Range object, then we could use `ActiveCell.Font.Size`.

####  Cell Content Manipulation

The `Value` property represents the contents of a cell. This can be used to change the contents of a cell.

```VBA
Range("B3").Value = 19  ' Set the value of cell B3 to 19

Range("D9").Value = "Text"  ' Set the value of cell D9 to the string "Text"
```

Of course, we can use dot notation to change the contents of cells in more specific locations.

```VBA
' This statement sets the value of cell B3 in Sheet2 of another open
' workbook to 19:
Workbook("Book5.xlsx").Sheets("Sheet2").Range("B3").Value = 19
```

Note that not using the `Value` property would have the same effect, since if no other property is specified, then the value of the cell is modified by default.

```VBA
Range("B3") = 19  ' Set the value of cell B3 to 19
```

Of course, it is also possible to change the value (or any other property, such as font size) of a cell based on another cell, or even itself.

```VBA

Range("B3") = Range("Al")  ' Set the value of cell B3 to the value of cell A1

Range("D2") = Range("D2") + 1  ' Increase the value of cell D2 by 1
```

The `ClearContents` method can be used to erase the contents of a cell.

```VBA
Range("B3").ClearContents  ' Erase the contents of cell B3
```

#### Text Formatting

To format text we will access the `Font` property. Within the IDE simply typing `Range("A1").Font.` will reveal a list of properties belonging to the `Font` property. For reference, we will list some of these below.

The `Size` property can be used to change the text size.

```VBA
Range("B3").Font.Size = 18  ' Format cell B3's contents to font size 18
```

The `Bold`, `Italic`, and `Underline` properties can be used to give text the bold, italic, and underline emphasis, respectively.

```VBA
Range("B3").Font.Bold = True  ' Bold the contents of cell B3

Range("D6").Font.Italics = True  ' Italicize the contents of cell D6

Range("A1").Font.Underline = True  ' Underline the contents of cell A1
```

The `Name` property can be used to set the font style.

```VBA
Range("B3").Font.Name = "Arial"  ' Format cell B3's contents to Arial font
```

#### Borders
Here are a couple of other commonly used properties.

The `Borders` property can be used to add a border to cells. Similar to the `Font` property, we can use the IDE to reveal a list of properties belonging to the `Borders` property.

```VBA
Range("B3:B9").Borders.Value = 1  ' Add a border to cells B3 to B9

Range("B3:B9").Borders.Weight = 4  ' Make the borders as thick as possible

Range("A2").Borders.Value = 0  ' Remove borders from cell A2
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

#### The `ColorIndex` Property

Unfortunately, the `ColorIndex` property is limited to only 56 colours, and depends on the colour theme of the application. The colour-index values for the default colour theme can be seen [here](https://docs.microsoft.com/en-us/office/vba/api/excel.colorindex#remarks).

```VBA
' The following statement sets the colour of the B3 cell to blue, and
' the font to white:
Range("B3").Interior.ColorIndex = 5
Range("B3").Font.ColorIndex = 2

Range("B3:B9").Border.ColorIndex = 4  ' Add a green border to cells B3 to B9

Sheets("Sheet7").Tab.ColorIndex = 45  ' Set the colour of tab Sheet7 to orange
```

#### The `Color` Property

Using the `Color` property is similar to using the `ColorIndex` property, except it uses RGB colour codes. Attempting to use this on versions of Excel older than 2007 will result in an approximate colour being chosen from the colour palette of 56 colours.

```VBA
' The following statement sets the colour of the B3 cell to blue, and
' the font to white:
Range("B3").Interior.Color = RGB(0, 0, 255)
Range("B3").Font.Color = RGB(255, 255, 255)

' This statement adds a green border to cells B3 to B9.
Range("B3:B9").Border.Color = RGB(0, 255, 0)

' This statement sets the colour of tab Sheet7 to orange.
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

Constants can't be changed later and are declared similarly to variables, but can only be Boolean, Byte, Integer, Long, Currency, Single, Double, Date, String, or Variant types. Since the value of a constant should already be known, we can specify the value and type in a single line.

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
VBA has four looping constructs - the `While`, `Do`, `For` and `For Each` loops.

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

### The `For Each` Loop

The basic syntax for a `For Each` loop is as follows:

```VBA
For Each item In collection
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
Sub subroutine1(variableName1 As Type1, variableName2 As Type2)
    statement  ' Can be dependent on variableName1 and variableName2
End Sub

Sub subroutine2()
    subroutine1 someArgument1, someArgument2
End Sub
```

By default, if a subroutine accepts arguments, then they are not optional - a subroutine called without the correct number of arguments will fail to execute. The `Optional` keyword can be used to specify optional arguments, but these must be specified after any non-optional arguments.

```VBA
Sub subroutine1(variableName1 As Type1, Optional variableName2 As Type2)
    statement ' Can be dependent on variableName1 and variableName2
End Sub

Sub subroutine2()
    ' All of the following subroutine calls are valid:
    subroutine1 someArgument1
    subroutine1 someArgument1, someArgument2
End Sub
```

### Passing Arguments by Value and by Reference

By default,  VBA is a pass-by-reference language. We now go into detail about what this means.

When passing arguments to a subroutine, we can choose to pass them either by value or by reference:
* Passing by reference: done by using the `ByRef` keyword. In VBA this is the default way to pass an argument if nothing is specified. In particular, this means that if a variable is passed as an argument, then its reference is transmitted.
* Passing by value: done by using the `ByVal` keyword. In particular, this means that if a variable is passed as an argument, then only its value is transmitted.

The differences are best demonstrated by example.

```VBA
' Below, two subroutines which square numbers have been created, but
' with different pass bys.

Sub calculateSquare1(ByRef number As Integer)
    number = number * number
End Sub

Sub calculateSquare2(ByVal number As Integer)
    number = number ^ 2
End Sub

' This subroutine tests the above two subroutines.
Sub test()
    Dim testNumber1 As Integer, testNumber2 As Integer
    testNumber1 = 20
    testNumber2 = 20

    calculateSquare1 testNumber1
    calculateSquare2 testNumber2

    Debug.Print(testNumber1)  ' Print: 400

    Debug.Print(testNumber2)  ' Print: 20
End Sub
```

Deciding between passing by value or passing by reference depends on both variable protection and performance. If we want to prevent a variable from being modified, then we should use `ByVal`. However, since passing by value works by copying the entire data contents of the variable, a large value type would be more efficiently passed by reference. See [here](https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/language-features/procedures/passing-arguments-by-value-and-by-reference) for more details.

### Functions

In VBA, subroutines and functions are similar, except functions must return a value. Using the function name within the function itself is used to determine the return value. Subroutines and functions can be called inside of each other, but function calls should have arguments enclosed in `( )`.

```VBA
' This function squares numbers.
Function calculateSquare(number As Integer)
    calculateSquare = number ^ 2
End Function

' This subroutine tests the calculateSquare function.
Sub test()
    result = calculateSquare(20)

    Debug.Print(result)  ' Print: 400
End Sub
```

Functions which are created inside modules can be used in a worksheet just like any other Excel function. Conversely, Excel functions can be used inside modules by using the WorksheetFunction object.

```VBA
WorksheetFunction.functionName
```

## Dialog Box Functions

Dialog boxes are useful for displaying information to the user, or for requesting user input. In this section we will discuss two commonly used dialog boxes, but many more exist, such as the [UserForm dialog box](#forms-and-controls) discussed later.

### The `MsgBox()` Function

The `MsgBox()` function can be used to display information to the user. The basic syntax is as follows:

```VBA
MsgBox("Hello World!")
```

The `MsgBox()` function can accept two other optional arguments to determine the buttons on the dialog box, and the title of the dialog box, i.e. `MsgBox(text, [buttons,] [title])`. See [here](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/msgbox-function#settings) for the full table on the settings of the buttons argument.

```VBA
' The following two lines produce the same dialog box:
MsgBox("Do you feel happy?", vbYesNo, "Mood Check")
MsgBox("Do you feel happy?", 4, "Mood Check")

' The following three lines produce the same dialog box:
MsgBox("Delete all contents?", vbYesNoCancel + vbExclamation, "Confirmation")
MsgBox("Delete all contents?", 3 + 48, "Confirmation")
MsgBox("Delete all contents?", 51, "Confirmation")
```

The `MsgBox()` function also returns values, which can be referred to using the following syntax:

```VBA
MsgBox(text, [buttons,] [title]) = returnValue
```

Either the constant name or the numerical value for the return value can be used. See [here](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/msgbox-function#return-values) for the full table of return values.

### The `InputBox()` Function

The `InputBox()` function can be used to request information from the user, which can then be stored in a variable. The basic syntax is as follows:

```VBA
InputBox("What is your name?")
```

The `InputBox()` function can accept two other optional arguments to determine the title of the dialog box and a default response, i.e. `MsgBox(text, [title,] [default])`.

## Events

In this section we will be using workbook or worksheet events (such as opening, closing and saving) to trigger VBA code. Throughout this section we will need to contain code in either the workbook or sheets modules. If code is placed into a different module type, then Excel will fail to find the code, and hence not execute it.

### Workbook Events

We can use special Private subroutine names to indicate which events will trigger VBA code. Note that the special subroutines in this section do not need to be Private, but should be for the sake of good code hygiene.

For workbook events, we need to ensure that the VBA code is contained in a workbook module. After ensuring that the first dropdown menu at the top of the IDE has "Workbook" selected, we can then use the second dropdown menu to see all possible workbook events. Selecting an event will automatically set up the special Private subroutine in which to enter code. Alternatively, we can also set up these special Private subroutines by simply entering the correct code.

We will now list some commonly used workbook events under which we can trigger VBA code, along with the associated subroutine syntax. See [here](https://docs.microsoft.com/en-us/office/vba/api/excel.workbook#events) for the full list of workbook events.

The `Open` event triggers code when the workbook is opened.

```VBA
Private Sub Workbook_Open()

End Sub
```

The `BeforeClose` event triggers code immediately before the workbook is closed.

```VBA
Private Sub Workbook_BeforeClose(Cancel As Boolean)
' The Cancel variable can be set to True to cancel the closing
' of the workbook.
End Sub
```

The `BeforeSave` event triggers code immediately before the workbook is saved.

```VBA
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
' The SaveAsUI variable returns True if the Save As dialog box
' was displayed.
'
' The Cancel variable can be set to True to cancel the saving
' of the workbook.
End Sub
```

The `AfterSave` event triggers code immediately after the workbook is saved.

```VBA
Private Sub Workbook_AfterSave(ByVal Success As Boolean)
' The Success variable returns True if the save operation
' was successful.
End Sub
```

The `BeforePrint` event triggers code immediately before anything in the workbook is printed.

```VBA
Private Sub Workbook_BeforePrint(Cancel As Boolean)
' The Cancel variable can be set to True to cancel the printing of
' anything in the workbook.
End Sub
```

The `SheetActivate` event triggers code each time the activated sheet is changed.

```VBA
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
' The Sh Worksheet object is the activated sheet.
End Sub
```

The `SheetBeforeDoubleClick` event triggers code immediately before a double-click on a cell in any worksheet.

```VBA
Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
' The Sh Worksheet object is the double-clicked worksheet.
'
' The Target Range object is the cell nearest to the mouse pointer when
' the double-click occurred.
'
' The Cancel variable can be set to True to cancel the default
' double-click action.
End Sub
```

The `SheetBeforeRightClick` event triggers code immediately before a right-click on a cell in any worksheet.

```VBA
Private Sub Workbook_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
' The Sh Worksheet object is the right-clicked worksheet.
'
' The Target Range object is the cell nearest to the mouse pointer when
' the right-click occurred.
'
' The Cancel variable can be set to True to cancel the default
' right-click action.
End Sub
```

The `SheetCalculate` event triggers code each time any worksheet's data is calculated or recalculated, or after any changed data is plotted on a chart.

```VBA
Private Sub Workbook_SheetCalculate(ByVal Sh As Object)
' The Sh Worksheet object is the changed sheet.
End Sub
```

The `SheetChange` event triggers code each time the contents of a cell in any worksheet are modified.

```VBA
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
' The Sh Worksheet object is the modified worksheet.
'
' The Target Range object is the changed range.
End Sub
```

The `SheetSelectionChange` event triggers code each time the selection changes on any worksheet.

```VBA
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
' The Sh Worksheet object is the worksheet containing the
' new selection.
'
' The Target Range object is the new selected range.
End Sub
```

The `NewSheet` event triggers code each time a new sheet is added to the workbook.

```VBA
Private Sub Workbook_NewSheet(ByVal Sh As Object)
' The Sh Worksheet object is the new sheet.
End Sub
```

The `SheetFollowHyperlink` event triggers code each time a hyperlink is clicked.

```VBA
Private Sub Workbook_SheetFollowHyperlink(ByVal Sh As Object, ByVal Target As Hyperlink)
' The Sh Worksheet object is the worksheet containing the hyperlink.
'
' The Target Hyperlink object is the destination of the hyperlink.
End Sub
```

### Worksheet Events

The logic for worksheet events is exactly the same as for workbook events, but applied at the worksheet level. Hence, all of the events in this section refer to occurrence in a single worksheet.

For worksheet events, we need to ensure that the VBA code is contained in a sheet module. Similar to the previous section, ensure that the first dropdown menu at the top of the IDE has "Worksheet" selected, and use the second dropdown menu to see all possible worksheet events.

Once again, we will now list some commonly used worksheet events under which we can trigger VBA code, along with the associated subroutine syntax. Some of the event names and syntaxes will be extremely similar to the previous section, but some of them are not exactly the same.

See [here](https://docs.microsoft.com/en-us/office/vba/excel/concepts/events-worksheetfunctions-shapes/worksheet-object-events) for the full list of worksheet events.

The `SelectionChange` event triggers code each time the selection changes in the worksheet.

```VBA
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
' The Target Range object is the new selected area.
End Sub
```

The `Activate` event triggers code each time the sheet is activated.

```VBA
Private Sub Worksheet_Activate()

End Sub
```

The `Deactivate` event triggers code each time the sheet is deactivated.

```VBA
Private Sub Worksheet_Deactivate()

End Sub
```

The `BeforeDoubleClick` event triggers code immediately before a double-click on a cell in the worksheet.

```VBA
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
' The Target Range object is the cell nearest to the mouse pointer
' when the double-click occurred.
'
' The Cancel variable can be set to True to cancel the default
' double-click action.
End Sub
```

The `BeforeRightClick` event triggers code immediately before a right-click on a cell in the worksheet.

```VBA
Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
' The Target Range object is the cell nearest to the mouse pointer when
' the right-click occurred.
'
' The Cancel variable can be set to True to cancel the default
' right-click action.
End Sub
```

The `Calculate` event triggers code each time the worksheet's data is calculated or recalculated.

```VBA
Private Sub Worksheet_Calculate()

End Sub
```

The `Change` event triggers code each time the contents of a cell in the worksheet are modified.

```VBA
Private Sub Worksheet_Change(ByVal Target As Range)
' The Target Range object is the changed range.
End Sub
```

The `FollowHyperlink` event triggers code each time a hyperlink is clicked.

```VBA
Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
' The Target Hyperlink object is the destination of the hyperlink.
End Sub
```

### Deactivating Events

It is possible that an event subroutine actually causes an event which triggers another event subroutine. It is easy to see that an undesirable infinite loop is a possibility.

The `Application.EnableEvents` property can be used to execute code without firing any events.

```VBA
Application.EnableEvents = False
statement
Application.EnableEvents = True
```

## Forms and Controls

UserForms are dialog boxes which allow for the addition of controls, such as the CheckBox, ComboBox and TextBox controls. With this, UserForms can be customised to a high degree.

Similar to workbook and worksheet events, we will be using events associated with UserForms, and their corresponding controls, to trigger VBA code. Throughout this section we will need to contain code in the UserForm module.

### UserForms

To insert a new UserForm module we can use the Project Explorer. Doing this will open the UserForm window from which we can open the following relevant windows:
- Properties window: allows for the modification of the UserForm's properties, such as appearance, behaviour and font. Open this via the right-click menu.
- Toolbox window: allows for the addition of controls. Opens automatically.
- IDE window: allows for VBA code triggered by events, to be attached to the UserForm. Open this by double-clicking the UserForm.

#### Events

Similar to the previous section, ensure that the first dropdown menu at the top of the IDE has "UserForm" selected, and use the second dropdown menu to see all possible UserForm events.

Typically UserForms are used because of the ability to use controls, but we will list a few commonly used UserForm events anyway.

The `Initialize` event triggers code when the UserForm is launched.

```VBA
Private Sub UserForm_Initialize()

End Sub
```

The `Click` event triggers code when the UserForm is clicked on.

```VBA
Private Sub UserForm_Click()

End Sub
```

The `Terminate` event triggers code when the UserForm is terminated.

```VBA
Private Sub UserForm_Terminate()

End Sub
```

#### Launching

We can of course use the IDE or a button in the sheet to launch a UserForm. We can also use the `Show` method to launch a UserForm from within a subroutine.

```VBA
Sub example()
    UserFormName.Show
End Sub
```

### Controls

All 14 available controls are available via the Toolbox window and can be placed anywhere in the UserForm. In this section we will only discuss the Label, TextBox, CommandButton, CheckBox and OptionButton controls. See [here](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/objects-microsoft-forms#controls) for more details on the rest of the controls.

We will also discuss some commonly used events and examples to add functionality to controls. As usual, all possible events can be viewed in the IDE via the second dropdown menu.

#### Label

The Label control displays text. The `Visible` property can be used to determine the visibility of the text. For example, it could be used to display an error message under certain conditions. Events are typically not associated with this control.

```VBA
labelName.Visible = True
```

#### TextBox

The TextBox control displays information entered by the user. The `Value` property can be used to access the information entered into the TextBox.

```VBA
Range("A1") = textboxName.Value  ' Store the value of textboxName in cell A1
```

The `Change` event triggers code each time the contents of the TextBox control is changed.

```VBA
Private Sub textboxName_Change()

End Sub
```

#### CommandButton

The CommandButton control starts, ends, or interrupts an action or series of actions. The `Click` event triggers code when the CommandButton control is clicked.

```VBA
Private Sub commandbuttonName_Click()

End Sub
```

The following example ties together the Label, TextBox and CommandButton controls:

```VBA
' This example is for a UserForm which asks the user to enter a
' numerical value.

' This subroutine sisplays labelError if the value in textboxNumerical
' is not a number.
Private Sub textboxNumerical_Change()
    If IsNumeric(textboxNumerical.Value) Then
        labelError.Visible = False
    Else
        labelError.Visible = True
    End If
End Sub

' This subroutine checks that the value entered into textboxNumerical
' is a number when buttonSubmit is clicked.  If the value is a number,
' then it is stored in the A1 cell.  Otherwise, a dialog box is shown.
Private Sub buttonSubmit_Click()
    If IsNumeric(textboxNumerical.Value) Then
        Range("A1") = textboxNumerical.Value
        Unload Me  ' Close the UserForm
    Else
        MsgBox("Incorrect value.")
    End If
End Sub
```

#### CheckBox

The CheckBox control displays the selection status of an item. The `Value` property can be used to access the status of the CheckBox control. The `Click` event triggers code when the CheckBox control is clicked.

The following example combines the `Value` property with the `Click` event:

```VBA
' This subroutine stores the status of checkboxExample in the A1 cell.
Private Sub checkboxExample_Click()
    If checkboxExample.Value = True Then
        Range("A1") = "Checked"
    Else
        Range("A1") = "Unchecked"
    End If
End Sub
```

#### OptionButton

The OptionButton control displays the status of one item in a group of choices. Similar to the CheckBox control, the `Value` property can be used to access the status of the OptionButton control, and the `Click` event triggers code when the OptionButton control is clicked.

The usage and syntax are similar, with the exception that only one OptionButton per group can be selected by the user. To create groups, the Frame control must first be placed before any OptionButton controls can be placed. Multiple Frame controls can be used to create multiple groups of OptionButton controls.

The `Controls` property of the Frame control returns the collection contained within the group. The following example makes use of this:

```VBA
' This subroutine enters text into a cell based on the user's choice.
Private Sub buttonConfirm_Click()
    Dim columnValue As String, rowValue As String

    ' This will iterate through each OptionButton in the Column group.
    For Each optionbuttonColumn In frameColumn.Controls
        If optionbuttonColumn.Value = True Then
            columnValue = optionbuttonColumn.Caption
        End If
    Next

    ' This will iterate through each OptionButton in the Row group.
    For Each optionbuttonRow In frameRow.Controls
        If optionbuttonRow.Value = True Then
            rowValue = optionbuttonRow.Caption
        End If
    Next

    Range(columnValue & rowValue) = "Cell chosen"
    Unload Me  ' Close the UserForm
End Sub
```
