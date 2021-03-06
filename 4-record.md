---
title: Record and Test Macros
nav: true
---
Please download <a href="media/IntroExcel-Part2-SampleWorkbook.xlsx" target="_blank">`IntroExcel-Part2-SampleWorkbook.xlsx`</a> to follow along with workshop activities.

**Throughout this workshop `Click` refers to a `Left Click` with your mouse**

# Record and Test Macros

If you frequently use Excel to work with similar worksheets and find yourself doing the same tasks again and again, `macros` can help you automate those tasks.

`Macros` allow you to record specific actions and complete those actions in other workbooks.

This lesson will show you how to all the steps needed to [create a macro](#create-a-macro):
1. [Enable the developer tab](#enable-the-developer-tab)
1. [Record a macro](#record-a-macro)
1. [Test a macro](#test-a-macro)
1. [View the Personal Macro Workbook](#view-the-personal-macro-workbook)
1. [Delete a macro](#delete-a-macro)

It will also share a few [helpful hints about macros](#helpful-hints-about-macros), offer resources for [more help](#more-help), detail the [difference between relative and absolute references](#difference-between-relative-and-absolute-references), and help you [find your Personal Macro Workbook file location](#find-your-personal-macro-workbook-file-location). 

## [Create a macro](#create-a-macro)

The `macro` tool is located under the `Developer` tab in Excel, which is hidden by default.

### [Enable the developer tab](#enable-the-developer-tab)

To enable the `Developer` tab on a PC:
* Click `File`
* Click `Options`
* Click `Customize Ribbon`
* Under `Main Tabs`, click to check the box next to `Developer`
* Click OK

To enable the `Developer` tab on a Mac:
* Click `Excel`
* Click `Preferences`
* Click `Ribbon & Toolbar`
* Under `Customize the Ribbon`, click `Main Tabs`, then click to check the box next to `Developer`
* Click Save

We can now see the `Developer` tab at the top of our Excel workbook.

### [Record a macro](#record-a-macro)

Let's navigate to `Sheet 3` in our workbook.

`Sheet 3` is an example of a worksheet with consistent organization (i.e. number of columns, column headings, etc.). 

Although the actual data changes each month (titles, authors, etc. are different), the organization stays exactly the same. Every time I download a new version of the worksheet, I complete the same formatting and organization tasks to make the sheet more readable. 

Since this Excel worksheet stays consistent, we can `record macros` that allow us to automate certain tasks when we download future versions of this worksheet.

We will focus on `recording a macro` for two types of tasks:
1. [Format a column](#format-a-column)
1. [Apply the PROPER function to a specific column](#apply-the-proper-function-to-a-specific-column)

#### [Format a column](#format-a-column)
* Click cell `A1`
* Navigate to the `Developer` tab and the `Code` section
* If it isn't already selected, click on `Use Relative References`
  * Learn more about the [difference between relative and absolute references](#difference-between-relative-and-absolute-references)
* Click `Record Macro`
* Under `Macro name`, type a name for the macro
  * Example: BookListFormatISBNColumnMacro
* In the `Store macro in` drop-down, select `Personal Macro Workbook`
  * This makes the `macro` available to use across all Excel workbooks
    * Learn how to [view the Personal Macro Workbook](#view-the-personal-macro-workbook)
* Enter a description
  * Example: Macro to format ISBN column to numbers in a book list.
* Click OK to start recording

Now we can record the `BookListFormatISBNColumnMacro` action:
* Click the `ISBN column` (or Column E) to highlight it
* Right click, then click `Format Cells`
* Under Category, click `Number`
* Set the `Decimal Places` to `0`
* Click OK
* Resize the column if necessary
* Click cell `A1`
* Navigate to the `Developer` tab and the `Code` section
* Click `Stop Recording`

#### [Apply the PROPER function to a specific column](#apply-the-proper-function-to-a-specific-column)
This macro will include four parts:
* [Part 1: Insert a new column](#part-1-insert-a-new-column)
* [Part 2: Insert the PROPER function](#part-2-insert-the-proper-function)
* [Part 3: Copy and paste values](#part-3-copy-and-paste-values)
* [Part 4: Delete a column](#part-4-delete-a-column)

##### [Part 1: Insert a new column](#part-1-insert-a-new-column)
* Click cell `A1`
* Navigate to the `Developer` tab and the `Code` section
* If it isn't already selected, click `Use Relative References`
  * Learn more about the [difference between relative and absolute references](#difference-between-relative-and-absolute-references)
* Click `Record Macro`
* Under `Macro name`, type a name for the macro
  * Example: BookListInsertColumnMacro
* In the `Store macro in` drop-down, select `Personal Macro Workbook`
  * This makes the `macro` available to use across all Excel workbooks
* Enter a description
  * Example: Macro to insert a new column in a book list.
* Click OK to start recording

Now we can record the `BookListInsertColumnMacro` action:
* Click the `Author column` (or Column B) to highlight it
* Right click, then click `Insert`
* Click cell `A1`
* Navigate to the `Developer` tab and the `Code` section
* Click `Stop Recording`

##### [Part 2: Insert the PROPER function](#part-2-insert-the-proper-function)
* Click cell `A1`
* Navigate to the `Developer` tab and the `Code` section
* If it isn't already selected, click on `Use Relative References`
  * Learn more about the [difference between relative and absolute references](#difference-between-relative-and-absolute-references)
* Click `Record Macro`
* Under `Macro name`, type a name for the macro
  * Example: BookListProperFunctionMacro
* In the `Store macro in` drop-down, select `Personal Macro Workbook`
  * This makes the `macro` available to use across all Excel workbooks
* Enter a description
  * Example: Macro to insert the PROPER function in a whole column in a book list.
* Click OK to start recording

Now we can record the `BookListProperFunctionMacro` action:
* Click `Cell B1`
* Type `=PROPER(A1)`
  * Option 1: Type `A1` **OR**
  * Option 2: Click `A1`
* Hit `Enter` on your keyboard
* Click `Column B` to highlight it
* Navigate to the `Home` tab and the `Editing` section
* Click `Fill`
* Click `Down` to copy the forumula to all relevant cells
* Click cell `A1`
* Navigate to the `Developer` tab and the `Code` section
* Click `Stop Recording`

##### [Part 3: Copy and paste values](#part-3-copy-and-paste-values)
* Click cell `A1`
* Navigate to the `Developer` tab and the `Code` section
* If it isn't already selected, click `Use Relative References`
  * Learn more about the [difference between relative and absolute references](#difference-between-relative-and-absolute-references)
* Click `Record Macro`
* Under `Macro name`, type a name for the macro
  * Example: BookListCopyPasteValuesMacro
* In the `Store macro in` drop-down, select `Personal Macro Workbook`
  * This makes the `macro` available to use across all Excel workbooks
* Enter a description
  * Example: Macro to copy and paste values in a book list.
* Click OK to start recording

Now we can record the `BookListCopyPasteValuesMacro` action:
* Click `Column B` to highlight it
* Right click, then click `Copy`
* Click `Column A` to highlight it
* Right click `Column A` 
* Click `Paste Special`
* Under `Paste`, click `Values`
  * This will let us copy and paste just the text
* Click OK
* Click cell `A1`
* Navigate to the `Developer` tab and the `Code` section
* Click `Stop Recording`

##### [Part 4: Delete a column](#part-4-delete-a-column)
* Click cell `A1`
* Navigate to the `Developer` tab and the `Code` section
* Click `Use Relative References`
* Click `Record Macro`
* Under `Macro name`, type a name for the macro
  * Example: BookListDeleteDuplicateTitleColumnMacro
* In the `Store macro in` drop-down, select `Personal Macro Workbook`
  * This makes the `macro` available to use across all Excel workbooks
* Enter a description
  * Example: Macro to delete the duplicate title column in a book list.
* Click OK to start recording

Now we can record the `BookListDeleteDuplicateTitleColumnMacro` action:
* Click the duplicate `Title column` (or Column B) to highlight it
* Right click, then click `Delete`
* Click cell `A1`
* Navigate to the `Developer` tab and the `Code` section
* Click `Stop Recording`

### [Test a macro](#test-a-macro)
Now that we have recorded these `macros`, let's test them on a new version of the same data.

First, let's download <a href="images/IntroExcel-Part2-SampleTestaMacroWorkbook.xlsx" target="_blank">`IntroExcel-Part2-SampleTestaMacroWorkbook.xlsx`</a>, open it on your computer, and click `enable editing` if prompted.

* Click cell `A1`
* Navigate to the `Developer` tab and the `Code` section
* Click `Macros`
* Click to highlight the `BookListDeleteColumnMacro`
* Click `Run`

We can then use the same steps to test the other macros:
* BookListFormatISBNColumnMacro
* BookListInsertColumnMacro
* BookListProperFunctionMacro
* BookListCopyPasteValuesMacro
* BookListDeleteColumnMacro

### [View the Personal Macro Workbook](#view-the-personal-macro-workbook)

Since we stored all the `macros` in the `Personal Macro Workbook`, we can `run` these in any Excel workbook. When you open Excel, the `Personal Macro Workbook` is hidden by default and you must `unhide` this workbook to delete a `macro`. 

To view the `Personal Macro Workbook`
* Open an Excel workbook
* Navigate to the `View` tab and the `Window` section (on a Mac, navigate to the `Window` tab) 
* Click `Unhide`
* In the `Unhide` window, click `Personal`
* Click OK

To hide the `Personal Macro Workbook`
* Navigate to the `Personal Macro Workbook`
* Navigate to the `View` tab and the `Window` section (on a Mac, navigate to the `Window` tab) 
* Click `Hide`

When you `record macros` or hide/unhide the `Personal Macro Workbook`, you must save your changes to `PERSONAL.XLSB` when prompted. You will see the following prompt when you attempt to close Excel after making changes to the `Personal Macro Workbook`:

![Microsoft Excel Personal Macro Workbook Save Prompt](media/PersonalMacroWorkbook_SavePrompt.JPG "Microsoft Excel Personal Macro Workbook Save Prompt")

To save your changes, click `Save`.

### [Delete a macro](#delete-a-macro)
After following the instructions above to [view the Personal Macro Workbook](#view-the-personal-macro-workbook):
* Navigate to the `Developer` tab and the `Code` section
* Click `Macros`
* Click to highlight the macro you want to `delete`
  * Example: BookListDeleteColumnMacro
* Click `Delete`
* When prompted: `Do you want to delete macro BookListDeleteColumnMacro?`, click `Yes`

When you `delete a macro`, you must save your changes to `PERSONAL.XLSB` when prompted.

### [Helpful hints about macros](#helpful-hints-about-macros)
* Record `separate macros` for each task
* Use `relative references`
  * Learn more about the [difference between relative and absolute references](#difference-between-relative-and-absolute-references)
* Click cell `A1` before you `start and stop` recording to keep `relative references` consistent
* Create `unique` macro names that are recognizable
  * Add a description to the description field

## [More Help](#more-help)

To learn more about `macros`, visit <a href="https://support.office.com/en-us/article/automate-tasks-with-the-macro-recorder-974ef220-f716-4e01-b015-3ea70e64937b" target="_blank">Microsoft's Macro website</a>.

### [Difference between relative and absolute references](#difference-between-absolute-and-relative-references)
This lesson has focused on using `relative references` when recording macros as it gives you more flexibility. 
* Recording a macro with `relative references` allows you to `run the macro` anywhere in a workbook 
  * This is why it is important to `start and stop` recording on cell `A1` when using `relative references`

In contrast, recording a macro with `absolute references` allows you to only `run the macro` on a specific cell or cell range, column, or row.
* Using `absolute references` can be useful when you plan to `run the macro` in the exact same spot across several workbooks
  * No matter where you `start and stop` your recording, the macro will always run on the same location in your worksbook.
  
### [Find your Personal Macro Workbook file location](#find-your-personal-macro-workbook-file-location)
* Open `File Explorer` (or `Finder` on a Mac)
* Navigate to `Local Disk (C:)`
* Double-click `Users`
* Double-click your `username`
* At the top of the `File Explorer` window, navigate to the `View` tab and the `Show/hide` section
* Check the box for `Hidden items`
* Double-click `AppData`
* Double-click `Roaming`
* Double-click `Microsoft`
* Double-click `Excel`
* Double-click `XLSTART`
* Navigate to your `Personal Macro Workbook` (file name = PERSONAL.XLSB)
