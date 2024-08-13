## Step 1: Install Python and the necessary libraries

Open the terminal in visual studio by pressing ctrl + j and run the following command to install required libraries:
pip install pandas matplotlib tk numpy logging openpyxl scipy

## Step 2: Create a folder

Create a new folder and place all the Python scripts within this folder.

## Step 3: Setup project folder

Create a project folder wherever you choose. This folder will contain the data files you need.

## Step 4: Add data files

Place the image_luminance.csv file and the excel echem file for the test within the project folder you created.

## Step 5: Run InterpolateData.py

In Visual Studio Code:
- Press the run button at the top right on the InterpolateData.py script, or
- Press Ctrl + J to open the terminal and type:
python InterpolateData.py
This will start the GUI.

## Step 6: Using the GUI

- In the GUI window that appears, press Select Directory and select the project folder containing the two data files.
- Adjust smoothing parameters as needed. For 0.5 min capture intervals, the default parameters should work well. For other intervals, adjust the smoothing parameters to avoid losing data granularity. The smoothing feature uses a Savitzky-Golay filter with a 2nd order polynomial fitting function—this can be adjusted within the script if needed.
- Click Combine Data and wait 5-10 seconds for the process to complete.

### Info

This process will:
1. Extract all the echem data from the excel sheet and create a new CSV file named Echem_extract.csv.
2. Interpolate values by taking the timestamp for the brightness measurement, finding the two closest points in the echem data, and creating a linear estimation between these points for the brightness timestamp voltage and current values.
3. Smooth the voltage, current, and brightness columns with respect to test time.
4. Take the derivative of the brightness with respect to time and smooth this derivative using the Savitzky-Golay filter.
5. Append all these (smoothed) data columns to a new CSV file named combined_data.csv within the project folder.

## Step 7: Check for errors

If no errors occur, you can proceed.

## Step 8: Create and view graph

- Click Create Graph, which will open another GUI (possibly behind the current window if not visible).
- Alternatively, run the GraphBrightnessData.py script:
python GraphBrightnessData.py
- In the GUI, click Select File and choose the newly created combined_data.csv file in your project folder.
- Select which (sequential) cycles you'd like to graph (by default, all are selected).
- Use the checkboxes to select which y-axes you'd like to graph. All data will be plotted against time on the x-axis. You can also copy the CSV data to OriginPro for further graphing.

### Note

This script graphs the smoothed versions of each column. To graph the non-smoothed version or to change the smoothing level, repeat the data interpolation process with the new parameters. This will overwrite the combined_data.csv file with your new settings.

## Step 9: Save created plot

- The created plot will be saved in a new folder within the project folder named Graphs.
- The graph is saved as a PNG file with transparency, which is useful for documents or PowerPoint presentations, but this can be changed by modifying a few lines in the Python script.
- The pop-up window is for quick analysis but may lack formatting by default. Use the configuration menu to adjust settings.
- If an axis for one of the datasets is missing, expand the space on the side of the plot. You can zoom in, move the data, and use the save button at the bottom to save the plot image.
