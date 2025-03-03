# Excel-Data-Modifier for Oil & Gas Production

**Summary:**
- Allows user to select an excel file and create a chart based on observed oil and gas production profiles and modify the production data by clicking and dragging the discrete data points.
  - The user can save the new data to a different sheet within the same excel file.

- **Purpose:**
- Unfortunately, in the oil and gas industry, production data can be erronous.
- If the errors are obvious but, not fixed during training, these can impact model performance.
- This GUI allows editing large oil & gas production data so that the user can train more accurate machine learning or deep learning models (e.g. ANN, Recurrent Neural Networks, LSTM, GRU, etc.).

**This project was done with common oil & gas production data in the common format:***
- A column: ID's
- C column: Dates
- D-E-F and H-I-J: Datas that can be modified.
- G: Number of days in a month.
- Other columns: These columns are generally irrelevant and was not considered in this GUI.


**How to install:**
- Create a new environment by running this line of code. "conda create --name myenv --file excel-data-modifier-env.txt". Here "myenv" be any name you like. 
  Run the code.

**How to use:**
- Select File and then Open Excel
- Select the excel file you want to work with
- You can select a sheet. Sheet 1 will be selected by default. Sheet 1 is the name of the sheet, and it can not be overriden.
- Select an id of the data you want.
- Whenever you make a selection, you must click the confirm button below.
- Select confirm selection button to draw a chart.
- The app will draw a chart of all the rows by column that has same id.
- Column D is active by default.

- The names of columns are as follows ["Montly Oil", "Montly Gas", "Montly Water", "Daily Oil", "Daily Gas", "Daily Water"] for [D, E, F, H, I, J]
  - These names were chosen for this specific oil and gas project. Please feel free to modify any part of the project as you see fit.

![image](https://user-images.githubusercontent.com/33734353/229098040-85ecde0d-a1ec-4db7-be7f-085937079a52.png)
  
**On right side, there are 2 listboxes:**
- Top listbox shows names of the columns that has data. User can double click on it and select **Add/Remove Graph** button to add or remove graph from the chart. You can see multiple charts on this graph at once.
- Bottom List box shows the datas of the selected chart. By selected, it means the item that was double clicked on top list box.
- Checkbox on right side, allows the days (column G) to be added/removed to the graph.

**On left side there are 5 buttons:**
- **Update selections**: Allows user to select different sheet and id from the same excel file.
- **Move Y point**: Allows user to modify data by selecting on graph and dragging it vertically. It must be selected chart. Press esc to cancel.
- **Delete data before**: By clicking the graph point, it deletes all the data before the clicked chart. Including other related datas to it. Press esc to cancel.
- **Delete data after**: Same as above, but all the data after the clicked chart. Press esc to cancel.
- **Save data**: Allows user to save the modified data to a new sheet, override existing sheet. "Sheet 1" can't be overriden. There should be at least one sheet stays original.

Below is an example of an multiphase oil & gas well production predicted using ANN & RNN (LSTM, GRU, Bi-LSTM, C-Bi-LSTM) models. In this case, the same model can predict oil, gas, and water simultanously.


![Picture29](https://github.com/user-attachments/assets/26a859a7-11d2-46c7-99fd-1be0ba346e64)
![Picture28](https://github.com/user-attachments/assets/d1eef3bf-fe5b-4018-832a-65972158c186)
![Picture27](https://github.com/user-attachments/assets/517177e2-838d-4831-a7dc-6ad50333a67d)

- The journal paper associated with the images above can be found in the link below. If you use this work, please cite the journal paper in this link:
- https://www.sciencedirect.com/science/article/abs/pii/S2949891024000587
