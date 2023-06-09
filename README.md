# Excel-Data-Modifier
Allows user to select an excel file and create a chart, then modify the chart by clicking and dragging the data. Then user can save the new data to the same excel file, but different sheet.

This project was done for specific type of excel data in mind.
A column: ID's
C column: Dates
D-E-F and H-I-J: Datas that can be modified.
G: Number of days in a month.
Other columns: They are irrelevant.

Purpose: Simplify big data, or modify it using this app, so the user can run prediction algorithm. Excel removed the ability to drag and move chart data, so it seemed like this app was necessary.

How to intall:
  Create a new environment by running this line of code. "conda create --name myenv --file excel-data-modifier-env.txt". Here "myenv" be any name you like. 
  Run the code.
How to use:
  Select File and then Open Excel
  Select the excel file you want to work with
  You can select a sheet. Sheet 1 will be selected by default. Sheet 1 is the name of the sheet, and it can not be overriden.
  Select an id of the data you want.
  Whenever you make a selection, you must click the confirm button below.
  Select confirm selection button to draw a chart.
  The app will draw a chart of all the rows by column that has same id.
  Column D is active by default.

  The names of columns are as follows ["Montly Oil", "Montly Gas", "Montly Water", "Daily Oil", "Daily Gas", "Daily Water"] for [D, E, F, H, I, J]
    It was done for this specific project. Feel free to modify any part of the project as you see fit.
   ![image](https://user-images.githubusercontent.com/33734353/229098040-85ecde0d-a1ec-4db7-be7f-085937079a52.png)
  
  On right side, there are 2 listboxes. Top listbox shows names of the columns that has data. User can double click on it and select "Add/Remove Graph" button to add or remove graph from the chart. You can see multiple charts on this graph at once.
    Bottom List box shows the datas of the selected chart. By selected, it means the item that was double clicked on top list box.
    Checkbox on right side, allows the days (column G) to be added/removed to the graph as secondary chart.
    
  On left side there are 5 buttons.
    Update selections button: Allows user to select different sheet and id from the same excel file.
    Move Y point button: Allows user to modify data by selecting on graph and dragging it vertically. It must be selected chart. Press esc to cancel.
    Delete data before: By clicking the graph point, it deletes all the data before the clicked chart. Including other related datas to it. Press esc to cancel.
    Delete data after: Same as above, but all the data after the clicked chart. Press esc to cancel.
    Save data: Allows user to save the modified data to a new sheet, override existing sheet. "Sheet 1" can't be overriden. There should be at least one sheet stays original.
