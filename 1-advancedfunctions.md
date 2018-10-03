---
title: Insert Advanced Functions
nav: true
---
Please download <a href="images/IntroExcel-Part2-SampleWorkbook.xlsx" target="_blank">`IntroExcel-Part2-SampleWorkbook.xlsx`</a> to follow along with workshop activities.

**Throughout this workshop `Click` refers to a `Left Click` with your mouse**

# Insert Advanced Functions

Excel offers many other functions that you can use to work with your data. These include [math and trig](#math-and-trig), [statistical](#statistical), and [text](#text) functions.

For a refresher on inserting functions, visit the <a href="https://jylisadoney.github.io/intro-excel-1/4-insert.html" target="_blank">Introduction to Excel Part 1: Insert Functions</a> page.

## [Math and Trig](#math-and-trig)
* Difference (no function exists)
  * Click on the cell where you want to calculate the difference 
  * Type: =CellA-CellB
    * Example: =A2-A3
* Divide (no function exists)
  * Click on the cell where you want to calculate the quotient 
  * Type: =CellA/CellB
    * Example: =A2/A3
* <a href="https://support.office.com/en-us/article/product-function-8e6b5b24-90ee-4650-aeec-80982a0512ce" target="_blank">PRODUCT</a>: multiply cells
* <a href="https://support.office.com/en-us/article/sum-function-043e1c7d-7726-4e80-8f32-07b23e057f89" target="_blank">SUM</a>: add cells

Learn more about <a href="https://support.office.com/en-us/article/math-and-trigonometry-functions-reference-ee158fd6-33be-42c9-9ae5-d635c3ae8c16" target="_blank">Math and Trig functions</a>.

## [Statistical](#statistical)
* <a href="https://support.office.com/en-us/article/average-function-047bac88-d466-426c-a32b-8f33eb960cf6" target="_blank">AVERAGE</a>: identify the average/mean for a range of cells
* <a href="https://support.office.com/en-us/article/countif-function-e0de10c6-f885-4e71-abb4-1f464816df34" target="_blank">COUNTIF</a>: count the number of cells in a range that meet specific criteria 
* <a href="https://support.office.com/en-us/article/max-function-e0012414-9ac8-4b34-9a47-73e662c08098" target="_blank">MAX</a>: identify the largest number contained in a range of cells
* <a href="https://support.office.com/en-us/article/median-function-d0916313-4753-414c-8537-ce85bdd967d2" target="_blank">MEDIAN</a>: identify the median for a range of cells
* <a href="https://support.office.com/en-us/article/min-function-61635d12-920f-4ce2-a70f-96f202dcc152" target="_blank">MIN</a>: identify the smallest number contained in a range of cells
* <a href="https://support.office.com/en-us/article/mode-function-e45192ce-9122-4980-82ed-4bdc34973120" target="_blank">MODE</a>: identify the most frequently occuring value for a range of cells
* <a href="https://support.office.com/en-us/article/stdev-function-51fecaaa-231e-4bbb-9230-33650a72c9b0" target="_blank">STDEV</a>: identify the standard deviation for a range of cells

Learn more about <a href="https://support.office.com/en-us/article/statistical-functions-reference-624dac86-a375-4435-bc25-76d659719ffd" target="_blank">Statistical functions</a>.

## [Text](#text)
* <a href="https://support.office.com/en-us/article/concatenate-function-8f8ae884-2ca8-4f7a-b093-75d702bea31d" target="_blank">CONCATENATE</a>: merge data from two cells into one cell
* <a href="https://support.office.com/en-us/article/lower-function-3f21df02-a80c-44b2-afaf-81358f9fdeb4" target="_blank">LOWER</a>: convert text to lowercase
* <a href="https://support.office.com/en-us/article/proper-function-52a5a283-e8b2-49be-8506-b2887b889f94" target="_blank">PROPER</a>: capitalize the first letter in each word
* <a href="https://support.office.com/en-us/article/upper-function-c11f29b3-d1a3-4537-8df6-04d0049963d6" target="_blank">UPPER</a>: convert text to uppercase

Learn more about <a href="https://support.office.com/en-us/article/text-functions-reference-cccd86ad-547d-4ea9-a065-7bb697c2a56e" target="_blank">Text functions</a>.

### Practice the CONCATENATE Function
Let's navigate to `Sheet 1` and practice the `CONCATENATE` function on the `Salesperson` column.

Right now, the data in the `Salesperson` column appears as `last-name, first-name`. We can use the `CONCATENATE` function to make the worksheet more visually appealing by changing the data in this column to `first-name last-name`

This requires six steps:
1. [Insert a new column](#insert-a-new-column)
1. [Split text into multiple columns](#split-text-into-multiple-columns)
1. [Insert another new column](#insert-another-new-column)
1. [Insert CONCATENATE function](#insert-concatenate-function)
1. [Replace old data](#replace-old-data)
1. [Delete duplicate data](#delete-duplicate-data)

#### [Insert a new column](#insert-a-new-column)
* Click on `Column C` (or the Sales Q1 column) to highlight it
* Right click then click on `Insert` to add a column

#### [Split text into multiple columns](#split-text-into-multiple-columns)
* Navigate to the `Data` tab and `Data Tools` section
* Click on `Column B` (or the Salesperson column) to highlight it
* Click on `Text to Columns`
* Select the `Delimited` option in the pop-up window
* Click `Next`
* Click the check-mark next to `Tab` to uncheck the box
* Click to check the box for `Comma` then check the box for `Space`
  * We can now see a preview of what our data will look like
* Click `Next`
* Click `Finish`

#### [Insert another new column](#insert-another-new-column)
* Click on `Column D` (or the Sales Q1 column) to highlight it
* Right click then click on `Insert` to add a column

#### [Insert CONCATENATE function](#insert-concatenate-function)
* Click on `Cell D2`
* Type `=CONCATENATE(D2," ",B2)`
  * Be sure to place a `space` between the quotation marks in this function
* Hit `Enter` on your keyboard
* Click on `Cell D2`
* Hold down the `Shift` key on your keyboard
* Scroll to and click on `Cell D23` to highlight the entire data range
* Hit `Ctrl + D` on your keyboard (or `Cmd + D` on a Mac) to copy the function to all the highligted cells

#### [Replace old data](#replace-old-data)
* Click on `Cell D2` to highlight it
* Hold down the `Shift` key on your keyboard
* Scroll to and click on `Cell D23` to highlight the entire data range
* Right click on the highlight data, then click on `Copy`
  * You can also hit `Ctrl + C` (or `Cmd + C` on a Mac) on your keyboard to copy data
* Click on `Cell B2` (in the Salesperson column) to highlight it
* Right click on `Cell B2` 
* Click on `Paste Special`
* Under `Paste`, click on `Values`
  * This will let us copy and paste just the text

#### [Delete duplicate data](#delete-duplicate-data)
* Click on `Column C` to highlight it
* Hold down the `Shift` key on your keyboard and click on `Column D` to highlight it too
* Right click, then click on `Delete`

## More Help
To learn more about `functions`, open Excel and navigate to the `Formulas` tab and the `Function Library` section. 

You can also visit <a href="https://support.office.com/en-us/article/excel-functions-by-category-5f91f4e9-7b42-46d2-9bd1-63f26a86c0eb?ui=en-US&rs=en-US&ad=US" target="_blank">Microsoft's Excel Functions (by Category) website</a> to learn about other functions.
