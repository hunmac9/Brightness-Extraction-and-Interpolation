import tkinter as tk
from tkinter import filedialog, messagebox
import os
import pandas as pd
import logging
from scipy.signal import savgol_filter
import subprocess

# Setting up logging
logging.basicConfig(filename='data_merger.log', level=logging.DEBUG, format='%(asctime)s:%(levelname)s:%(message)s')

def find_excel_file(input_dir):
    excel_files = [f for f in os.listdir(input_dir) if f.endswith(('.xlsx', '.xls'))]
    if len(excel_files) == 1:
        return os.path.join(input_dir, excel_files[0])
    elif len(excel_files) == 0:
        raise FileNotFoundError("No Excel files found in the directory.")
    else:
        raise ValueError("Multiple Excel files found in the directory.")

def ensure_echem_extract_exists(input_dir):
    echem_path = os.path.join(input_dir, 'Echem_Extract.csv')
    if not os.path.exists(echem_path):
        logging.warning(f"Echem_Extract.csv not found in directory {input_dir}. Looking for an Excel file to process.")
        try:
            excel_file_path = find_excel_file(input_dir)
            logging.info(f"Found Excel file: {excel_file_path}. Running EchemProcessing.py script.")
            result = subprocess.run(['python', 'EchemProcessing.py', excel_file_path], capture_output=True, text=True)
            if result.returncode != 0:
                logging.error(f"Failed to run EchemProcessing.py: {result.stderr}")
                messagebox.showerror("Processing Error", f"Failed to run EchemProcessing.py: {result.stderr}")
                return False
        except (FileNotFoundError, ValueError) as e:
            logging.error(str(e))
            messagebox.showerror("File Error", str(e))
            return False

        if not os.path.exists(echem_path):
            logging.error("Echem_Extract.csv still not found after running processing script.")
            messagebox.showerror("File Error", "Echem_Extract.csv not found in directory after running processing script.")
            return False
    logging.info("Echem_Extract.csv found or created successfully.")
    return True

def read_files(input_dir):
    try:
        if not ensure_echem_extract_exists(input_dir):
            return None, None

        echem_path = os.path.join(input_dir, 'Echem_Extract.csv')
        image_brightness_path = os.path.join(input_dir, 'image_luminance.csv')
        
        # Read data
        echem_data = pd.read_csv(echem_path)
        image_data = pd.read_csv(image_brightness_path)
        
        # Check if required columns are present
        if not {'Timestamp', 'Voltage(V)', 'Current(A)', 'Cycle_Index'}.issubset(echem_data.columns):
            raise ValueError("Echem_Extract.csv does not have the required columns.")
        if not {'Timestamp', 'Luminance'}.issubset(image_data.columns):
            raise ValueError("image_luminance.csv does not have the required columns.")
        
        logging.info("Files read successfully.")
        return echem_data, image_data
    except Exception as e:
        logging.error(f"Error in reading files: {e}")
        messagebox.showerror("File Error", f"An error occurred while reading the files: {e}")
        return None, None

def preprocess_data(echem_data, image_data):
    try:
        # Convert Timestamp to datetime
        echem_data['Timestamp'] = pd.to_datetime(echem_data['Timestamp'])
        image_data['Timestamp'] = pd.to_datetime(image_data['Timestamp'])

        # Normalize brightness data
        image_data['Luminance'] = image_data['Luminance'] * 100 / 255

        # Sort data by Timestamp
        echem_data.sort_values(by='Timestamp', inplace=True)
        image_data.sort_values(by='Timestamp', inplace=True)
        
        logging.info("Data preprocessing successful.")
        return echem_data, image_data
    except Exception as e:
        logging.error(f"Error in data preprocessing: {e}")
        messagebox.showerror("Data Error", f"An error occurred during data preprocessing: {e}")
        return None, None

def add_smoothed_column(data, num_points, column):
    try:
        smoothed_column_name = f"{column}_smooth"
        if num_points > 0 and num_points % 2 != 0:  # num_points must be odd for Savitzky-Golay filter
            data[smoothed_column_name] = savgol_filter(data[column], num_points, 2)
            logging.info(f"Smoothing applied on {column} with {num_points} points and stored in {smoothed_column_name}.")
        else:
            data[smoothed_column_name] = data[column]  # if smoothing is not applicable, copy the original data
            logging.warning(f"Smoothing points for {column} must be a positive odd number. No smoothing applied.")
        return data
    except Exception as e:
        logging.error(f"Error in smoothing {column}: {e}")
        raise

def convert_current_to_mA(echem_data):
    echem_data['Current(mA)'] = echem_data['Current(A)'] * 1000
    return echem_data

def combine_data(echem_data, image_data):
    try:
        # Initialize a list to hold combined data
        combined_data = []

        for _, row in image_data.iterrows():
            brightness_time = row['Timestamp']
            brightness_value = row['Luminance']
            brightness_smooth_value = row['Luminance_smooth']

            # Find the closest two voltage data points
            past_points = echem_data[echem_data['Timestamp'] <= brightness_time]
            future_points = echem_data[echem_data['Timestamp'] >= brightness_time]

            if (past_points.empty or future_points.empty) or (past_points.shape[0] < 2 or future_points.shape[0] < 2):
                # Skip if there are no valid previous or next points
                logging.warning(f"No valid interpolation points found for brightness timestamp {brightness_time}. Skipping this point.")
                continue

            previous_point = past_points.iloc[-1]
            next_point = future_points.iloc[0]

            # Linear interpolation
            t1 = previous_point['Timestamp']
            t2 = next_point['Timestamp']
            v1 = previous_point['Voltage(V)']
            v2 = next_point['Voltage(V)']
            i1 = previous_point['Current(mA)']
            i2 = next_point['Current(mA)']
            vi_smooth1 = previous_point['Voltage(V)_smooth']
            vi_smooth2 = next_point['Voltage(V)_smooth']
            ci_smooth1 = previous_point['Current(mA)_smooth']
            ci_smooth2 = next_point['Current(mA)_smooth']
            c_index = previous_point['Cycle_Index']
            
            if t1 == t2:
                voltage_interpolated = v1
                current_interpolated = i1
                voltage_smoothed_interpolated = vi_smooth1
                current_smoothed_interpolated = ci_smooth1
            else:
                # Calculate the interpolated values
                total_time = (t2 - t1).total_seconds()  # in seconds
                time_fraction = (brightness_time - t1).total_seconds() / total_time

                voltage_interpolated = v1 + (v2 - v1) * time_fraction
                current_interpolated = i1 + (i2 - i1) * time_fraction
                voltage_smoothed_interpolated = vi_smooth1 + (vi_smooth2 - vi_smooth1) * time_fraction
                current_smoothed_interpolated = ci_smooth1 + (ci_smooth2 - ci_smooth1) * time_fraction
            
            # Calculate test time in hours from the first combined data point
            test_time = (brightness_time - image_data['Timestamp'].min()).total_seconds() / 3600

            combined_data.append({
                'Timestamp': brightness_time,
                'Brightness': brightness_value,
                'Brightness_smooth': brightness_smooth_value,
                'Voltage(V)': voltage_interpolated,
                'Voltage(V)_smooth': voltage_smoothed_interpolated,
                'Current(mA)': current_interpolated,
                'Current(mA)_smooth': current_smoothed_interpolated,
                'Cycle_Index': c_index,
                'Test Time (h)': test_time
            })
        
        combined_df = pd.DataFrame(combined_data)

        # Add the derivative of brightness from the smoothed brightness data
        combined_df['Brightness Derivative'] = combined_df['Brightness_smooth'].diff() / combined_df['Test Time (h)'].diff()
        
        logging.info("Data combination successful.")
        return combined_df
    except Exception as e:
        logging.error(f"Error in combining data: {e}")
        messagebox.showerror("Data Error", f"An error occurred during data combination: {e}")
        return None

def save_combined_data(combined_df, output_dir):
    try:
        output_path = os.path.join(output_dir, 'combined_data.csv')
        combined_df.to_csv(output_path, index=False)
        logging.info(f"Combined data saved successfully to {output_path}.")
        messagebox.showinfo("Success", f"Combined data saved successfully to {output_path}.")
        return output_path
    except Exception as e:
        logging.error(f"Error in saving combined data: {e}")
        messagebox.showerror("Save Error", f"An error occurred while saving the combined data: {e}")

def select_directory(directory_label):
    input_dir = filedialog.askdirectory()
    if input_dir:
        logging.info(f"Selected directory: {input_dir}")
        directory_label.config(text=f"Selected Directory: {input_dir}")
    return input_dir

def combine_data_process(voltage_entry, current_entry, brightness_entry, brightness_derivative_entry, input_dir):
    if not input_dir:
        messagebox.showerror("Directory Error", "Please select a directory first.")
        return
    
    try:
        voltage_points = int(voltage_entry.get())
        current_points = int(current_entry.get())
        brightness_points = int(brightness_entry.get())
        brightness_derivative_points = int(brightness_derivative_entry.get())
    except ValueError:
        messagebox.showerror("Input Error", "Number of smoothing points must be an integer.")
        return
    
    if voltage_points < 0 or current_points < 0 or brightness_points < 0 or brightness_derivative_points < 0:
        messagebox.showerror("Input Error", "Number of smoothing points must be non-negative.")
        return

    try:
        echem_data, image_data = read_files(input_dir)
        if echem_data is None or image_data is None:
            return

        echem_data, image_data = preprocess_data(echem_data, image_data)
        if echem_data is None or image_data is None:
            return

        # Convert current to mA before smoothing
        echem_data = convert_current_to_mA(echem_data)

        # Smooth echem and image data as specified
        if voltage_points > 0:
            echem_data = add_smoothed_column(echem_data, voltage_points, 'Voltage(V)')
        if current_points > 0:
            echem_data = add_smoothed_column(echem_data, current_points, 'Current(mA)')

        if brightness_points > 0:
            image_data = add_smoothed_column(image_data, brightness_points, 'Luminance')

        # Combine data
        combined_df = combine_data(echem_data, image_data)
        if combined_df is None:
            return

        # Smooth the brightness derivative as specified
        if brightness_derivative_points > 0:
            combined_df = add_smoothed_column(combined_df, brightness_derivative_points, 'Brightness Derivative')

        # Save the combined and smoothed data
        combined_filepath = save_combined_data(combined_df, input_dir)
        return combined_filepath

    except Exception as e:
        logging.error(f"Error in data processing: {e}")
        messagebox.showerror("Processing Error", f"An error occurred during data processing: {e}")

def create_graph(combined_filepath):
    try:
        if combined_filepath and os.path.isfile(combined_filepath):
            subprocess.run(['python', 'GraphBrightnessData.py', combined_filepath], check=True)
            logging.info(f"Graph created successfully using {combined_filepath}.")
        else:
            raise FileNotFoundError(f"The file {combined_filepath} was not found.")
    except Exception as e:
        logging.error(f"Error in creating graph: {e}")
        messagebox.showerror("Graph Error", f"An error occurred while creating the graph: {e}")

def main():
    # Create the GUI window
    root = tk.Tk()
    root.title("Data Merger")

    input_dir = tk.StringVar()
    combined_filepath = tk.StringVar()
    
    label = tk.Label(root, text="Select the input directory containing the CSV files")
    label.pack(pady=5)

    directory_label = tk.Label(root, text="No directory selected")
    directory_label.pack(pady=5)

    select_button = tk.Button(root, text="Select Directory", command=lambda: input_dir.set(select_directory(directory_label)))
    select_button.pack(pady=10)

    # Frame for voltage smoothing input
    voltage_frame = tk.Frame(root)
    voltage_frame.pack(pady=5)
    voltage_label = tk.Label(voltage_frame, text="Voltage smoothing points (0 for no smoothing):")
    voltage_label.pack(side=tk.LEFT)
    voltage_entry = tk.Entry(voltage_frame)
    voltage_entry.pack(side=tk.LEFT)
    voltage_entry.insert(0, "21")
    
    # Frame for current smoothing input
    current_frame = tk.Frame(root)
    current_frame.pack(pady=5)
    current_label = tk.Label(current_frame, text="Current smoothing points (0 for no smoothing):")
    current_label.pack(side=tk.LEFT)
    current_entry = tk.Entry(current_frame)
    current_entry.pack(side=tk.LEFT)
    current_entry.insert(0, "21")
    
    # Frame for brightness smoothing input
    brightness_frame = tk.Frame(root)
    brightness_frame.pack(pady=5)
    brightness_label = tk.Label(brightness_frame, text="Brightness smoothing points (0 for no smoothing):")
    brightness_label.pack(side=tk.LEFT)
    brightness_entry = tk.Entry(brightness_frame)
    brightness_entry.pack(side=tk.LEFT)
    brightness_entry.insert(0, "101")
    
    # Frame for brightness derivative smoothing input
    brightness_derivative_frame = tk.Frame(root)
    brightness_derivative_frame.pack(pady=5)
    brightness_derivative_label = tk.Label(brightness_derivative_frame, text="Brightness derivative smoothing points (0 for no smoothing):")
    brightness_derivative_label.pack(side=tk.LEFT)
    brightness_derivative_entry = tk.Entry(brightness_derivative_frame)
    brightness_derivative_entry.pack(side=tk.LEFT)
    brightness_derivative_entry.insert(0, "41")
    
    combine_button = tk.Button(root, text="Combine Data", command=lambda: combined_filepath.set(combine_data_process(
        voltage_entry, current_entry, brightness_entry, brightness_derivative_entry, input_dir.get())))
    combine_button.pack(pady=10)
    
    create_graph_button = tk.Button(root, text="Create Graph", command=lambda: create_graph(combined_filepath.get()))
    create_graph_button.pack(pady=10)
    
    root.mainloop()

if __name__ == "__main__":
    main()
