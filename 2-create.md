---
title: Create Charts
nav: true
---
Please download <a href="images/IntroExcel-Part2-SampleWorkbook.xlsx" target="_blank">`IntroExcel-Part2-SampleWorkbook.xlsx`</a> to follow along with workshop activities.

**Throughout this workshop `Click` refers to a `Left Click` with your mouse**

# Create Charts and Graphs

Not only can Excel help you organize data and information, it can also help you create charts and graphs to visually share your data.

Today, we will [create a Bar Chart](#create-a-bar-chart) that compares each employee's sales in Q1 and Q2 and [describe the data](#describe-the-data) in more detail.

## [Create a bar chart](#create-a-bar-chart)

First, let's navigate to `Sheet1` in our workbook.
* Highlight all of our data
  * Option 1: Click `Cell A1`, hold down the `Shift` key on your keyboard, and click `Cell F23`
  * Option 2: Click `Cell A1` and drag to highlight all data **OR**
  * Option 3: Hit `Ctrl + A` your keyboard (`Cmd + A` for Mac users)

We can now work on creating our chart.
* Navigate to the `Insert` tab and the `Charts` section
* Click the picture of the `Bar Chart`
* Click `2-D Clustered Column` (option 1)

Right away, we see that this chart is a **mess** and the visualization is confusing -- but we can fix it!

### Edit the data in a bar chart

Since we only want to compare `sales in Q1 and Q2`, we aren't interested in `Total Sales` or `Hire Date`.

To remove `Total Sales` and `Hire Date` from our chart:
* Click the chart 
* Click the `Funnel Icon` (on a Mac, right click the chart and click `Select Data`)
* Under `Series`, click the check-mark next to `Total Sales` and `Hire Date` to uncheck the boxes (on a Mac, under `Series` click to highlight `Total Sales`, then click `Remove` and follow the same steps for `Hire Date`)
* Click `Apply` (on a Mac, click OK)

Our chart is looking much better - but including the state names in our chart makes it difficult to read.

To remove `State` names from our chart:
* Click the chart
* Click the `Funnel Icon` (on a Mac, right click the chart and click `Select Data`)
* Click `Names` (on a Mac, navigate to the `Category (X) axis labels box`)
* Under `Categories`, select the option for `Column B` (on a Mac, change the cell range to remove any reference to state names, for example: `=Sheet1!$B$2:$B$23`)
* Click `Apply` (on a Mac, click OK)

After these two small changes, our chart looks so much better!

## [Describe the data](#describe-the-data)
Another important part of creating visualizations is to add text that describes the data.

### Add a chart title
* Click the chart 
* Navigate to the `Design` tab and the `Chart Layouts` section
* Click `Add Chart Element`
* Hover over `Chart Title` and click `Above Chart`

We now have a text box that we can use to add a chart title.

To edit the text:
* Click the chart
* Click the `Chart Title` text box
* Type a title that describes the data
  * Maybe something like `Employee Sales by Quarter`

### Position the chart legend
* Click the chart
* Navigate to the `Design` tab and the `Chart Layouts` section
* Click `Add Chart Element`
* Hover over `Legend`
* Click `Right`

Our legend is now positioned to the `right` of our `bar chart`.

### Add a title to the horizontal axis (x-axis)
* Click the chart
* Navigate to the `Design` tab and the `Chart Layouts` section
* Click `Add Chart Element`
* Hover over `Axis Titles`
* Click `Primary Horizontal`

We now have a text box that we can use to describe this axis.

To edit the text:
* Click the chart
* Click the `Axis Title` text box
* Type a word or phrase that describes this axis
  * Maybe something like `Employee Names`

You can use the same steps to add a title to the `Primary Vertical axis`.

## More Help
To learn more about `charts and graphs`, open Excel and navigate to the `Insert` tab and the `Charts` section. 

You can also visit <a href="https://support.office.com/en-us/article/available-chart-types-in-office-a6187218-807e-4103-9e0a-27cdb19afb90?ui=en-US&rs=en-US&ad=US" target="_blank">Microsoft's Available Chart Types website</a>.
