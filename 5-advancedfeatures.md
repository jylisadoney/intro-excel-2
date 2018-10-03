---
title: Utilize Advanced Features
nav: true
---
Please download <a href="images/IntroExcel-Part2-SampleWorkbook.xlsx" target="_blank">`IntroExcel-Part2-SampleWorkbook.xlsx`</a> to follow along with workshop activities.

**Throughout this workshop `Click` refers to a `Left Click` with your mouse**

# Utilize Advanced Features

Excel offers many other features that you can use to work with your data.

These include [Templates](#templates) and [Data Validation](#data-validation). Links to [more help](#more-help) on each of these features is available at the bottom of the page.

## Templates

If you often use the same workbook layout, you can save it as a `template` and use it again in the future.

To use templates, we first have to `set the default personal templates location`.

### Set the default personal templates location
* Click on `File` then click on `Options`
* Click on the `Save` tab
* Navigate to the `Save workbooks` section
* Navigate to the `Deafult personal templates location` box
* Type this file path: `C:\Users\[UserName]\Documents\Custom Office Templates`
  * Replace [User Name] with your user name
* Click `OK`

Unless you want to change the `default personal templates location`, you will only need to set this once.

### Save a workbook as a template
* Open Excel
* Open the workbook you want to save as a template
* Click on `File` then click on `Export`
* Under `Export`, click on `Change File Type`
* Under the `Workbook File Types` section, double-click on `Template`
* Type the file name you want to use for the template
* Click on `Save` and close the template

### Create a workbook based on a template
* Open Excel
* Click on `File` then click on `New`
* Click on `Personal`
* Double-click on the template you want to use

Excel will then create a new workbook based on this template. Be sure to `Save` this new workbook.

## Data validation

Data validation in Excel controls `what` you can enter into a cell.

These first three steps are the same for any type of `data validation`.
* Click on the first cell where you want to `add data validation`
  * Do not select the `Header` cell
* Hold down the `Shift` key on your keyboard to select all cells where you want to add data validation
* Navigate to the `Data` tab and the `Data Tools` section
* Click on `Data Validation`

Follow the instructions below to learn how to:
* [Restrict data entry to whole numbers within limits](#restrict-data-entry-to-whole-numbers-within-limits)
* [Restrict data entry to a range of dates](#restrict-data-entry-to-a-range-of-dates)
* [Restrict data entry to a drop-down list of data values](#restrict-data-entry-to-a-drop-down-list-of-data-values)
* [Remove data validation](#remove-data-validation)
* [Specify the input message or error alert](#specify-the-input-message-or-error-alert)

### [Restrict data entry to whole numbers within limits](#restrict-data-entry-to-whole-numbers-within-limits)

Let's Navigate to `Sheet 1` to practice data validation.

* Click on `Cell G1` and rename it
  * Example Number of Mentees
* Click on `Cell G2`, hold down the `Shift` key on your keyboard, and click on `Cell G23` to highlight the cell(s) where you want to add data validation
* Navigate to the `Data` tab and the `Data Tools` section
* Click on `Data Validation`
* In the `Settings` tab:
  * Set the `Allow` drop-down box to `Whole number`
  * Set the `Data` drop-down box to `between`
  * Enter the allowed `Minimum` and `Maximum` range of whole numbers
    * For example, Minimum: 0 and Maximum: 5
* Click `OK`

Now, let's test our data validation
* Click on `Cell G2`
* Type the number `2`
* Hit `Enter` on your keyboard

Excel let us enter this value since it is a `whole number` between `0 and 5`.

If you try to enter a number **outside the specified range**, you will see an error message.
* Click on `Cell G3`
* Type the number `1.5`
* Hit `Enter` on your keyboard
* Click `Retry` to enter a value that meets the `data validation` requirements

{% include figure.html file="DataValidation_ErrorMessage.JPG" alt="Microsoft Excel Personal Macro Workbook Save Prompt" caption="Microsoft Excel Personal Macro Workbook Save Prompt" %}

### [Restrict data entry to a range of dates](#restrict-data-entry-to-a-range-of-dates)
* Click on the first cell where you want to `add data validation`
  * Do not select the `Header` cell
* Hold down the `Shift` key on your keyboard to select all cells where you want to add data validation
* Navigate to the `Data` tab and the `Data Tools` section
* Click on `Data Validation`
* In the `Settings` tab:
  * Set the `Allow` drop-down box to `Date`
  * Set the `Data` drop-down box to `greater than or equal to`
  * Enter the allowed `Start date`
    * For example, Start date: 01/01/1999
* Click `OK`

Now, if you try to enter a date **before** 01/01/1999, you will see an error message.

### [Restrict data entry to a drop-down list of data values](#restrict-data-entry-to-a-drop-down-list-of-data-values)
* Click on the first cell where you want to `add data validation`
  * Do not select the `Header` cell
* Hold down the `Shift` key on your keyboard to select all cells where you want to add data validation
* Navigate to the `Data` tab and the `Data Tools` section
* Click on `Data Validation`
* In the `Settings` tab:
  * Set the `Allow` drop-down box to `List`
  * In the `Source` box, type your data values and **separate them with commas without spaces**
    * For example, `Yes,No,Maybe`
* Click to check the `In-cell dropdown` box
* Click `OK`

Now, when you want to enter text, you can only select from the **allowed options**.

### [Remove data validation](#remove-data-validation)
* Click and drag to highlight the cell(s) where you want to remove data validation
* Navigate to the `Data` tab and the `Data Tools` section
* Click on `Data Validation`
* Click `Clear All`
* Click `OK`

### [Specify the input message or error alert](#specify-the-input-message-or-error-alert)
Excel also gives you the option to specify the `input message` or `error alert` to make is less generic.

#### Specify the input message
* Highlight the cells where you want to specify the `input message`
* Navigate to the `Data` tab and the `Data Tools` section
* Click on `Data Validation`
* Navigate to the `Input Message` tab
* Check the box next to `Show input message when cell is selected`
* Add a title and input message that describes the data validation requirements

#### Specify the error message
* Highlight the cells where you want to specify the `error alert`
* Navigate to the `Data` tab and the `Data Tools` section
* Click on `Data Validation`
* Navigate to the `Error Alert` tab
* Check the box next to `Show error alert after invalid data is entered`
* Under the `Style` drop-down box, select `Stop`, `Warning`, or `Information` to change the icon
* Add a title and error message that describes the data validation requirements

## [More Help](#more-help)
### Templates
To learn more about `templates`, visit <a href="https://support.office.com/en-us/article/save-a-workbook-as-a-template-58c6625a-2c0b-4446-9689-ad8baec39e1e" target="_blank">Microsoft's Save a Workbook as a Template website</a>.

### Data validation
To learn more about `data validation`, open Excel and navigate to the `Data` tab and the `Data Tools` section. 

You can also visit <a href="https://support.office.com/en-us/article/apply-data-validation-to-cells-29fecbcc-d1b9-42c1-9d76-eff3ce5f7249" target="_blank">Microsoft's Apply Data Validation to Cells website</a>.
