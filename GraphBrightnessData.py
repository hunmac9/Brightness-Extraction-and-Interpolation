import pandas as pd
import matplotlib.pyplot as plt
from tkinter import Tk, filedialog, Listbox, MULTIPLE, Label, Checkbutton, IntVar, Button, Scrollbar, END, StringVar
import os
from datetime import datetime

# Create the main Tkinter window
root = Tk()
root.title("Data Plotter")

# Variables
plotcurrent = IntVar(value=1)
plotbright = IntVar(value=1)
plotvolt = IntVar()
plotderiv = IntVar()
plotcycles = []
selected_file = StringVar()

def setup_plot_styles():
    # Set font properties
    plt.rcParams['font.family'] = 'Helvetica'
    plt.rcParams['font.weight'] = 'regular'  # Set to regular weight
    plt.rcParams['font.size'] = 16  # Suitable font size for readability
    plt.rcParams['axes.labelweight'] = 'regular'  # Regular weight for axis labels
    plt.rcParams['lines.linewidth'] = 1.5  # Line width
    plt.rcParams['xtick.labelsize'] = 14
    plt.rcParams['ytick.labelsize'] = 14
    plt.rcParams['xtick.major.size'] = 8
    plt.rcParams['ytick.major.size'] = 8
    plt.rcParams['xtick.major.width'] = 1.5
    plt.rcParams['ytick.major.width'] = 1.5

def customize_axis(ax, color, ylabel=None):
    if ylabel:
        ax.set_ylabel(ylabel, fontsize=16, color=color)
    ax.tick_params(axis='y', labelcolor=color, width=1.5, colors=color)
    ax.spines['right'].set_color(color)
    ax.spines['right'].set_linewidth(1.5)
    if ylabel:
        ax.yaxis.label.set_color(color)
    
    for tick in ax.yaxis.get_major_ticks():
        tick.label1.set_fontsize(14)
        tick.label1.set_fontfamily('Helvetica')

def adjust_plot_layout(fig, axes_list):
    # Adjust the positions of the right y-axes dynamically
    pos = 1  # Initial position of the first right y-axis
    delta_pos = 0.12  # Step between the axes

    for ax in axes_list:
        ax.spines['right'].set_position(('axes', pos))
        pos += delta_pos
        ax.spines['right'].set_visible(True)

    fig.tight_layout(pad=1.0)
    plt.subplots_adjust(left=0.167, right=0.85 + len(axes_list) * 0.07, top=0.967, bottom=0.2)

def main(filepath, plotcycles, plotcurrent, plotbright, plotvolt, plotderiv):
    setup_plot_styles()

    # Read data from CSV file
    data = pd.read_csv(filepath)
    # Sort data by time
    data = data.sort_values(by='Test Time (h)')

    cycle_data = data[data['Cycle_Index'].isin(plotcycles)]
    test_time = cycle_data['Test Time (h)']

    # Compute x-axis limits with margins
    x_min = 0
    x_max = test_time.max() * 1.02  # 2% margin on right

    fig, ax1 = plt.subplots(figsize=(8, 6))  # Adjusted size for better space management
    ax1.set_xlim(x_min, x_max)

    axes_list = []
    if plotcurrent.get() == 1:
        current = cycle_data['Current(mA)_smooth']
        color = 'black'
        ax1.set_xlabel('Test Time (h)', fontsize=16)
        ax1.set_ylabel('Current (mA)', fontsize=16, color=color)
        ax1.plot(test_time, current, color=color, linewidth=1.5)
        ax1.tick_params(axis='x', width=1.5, colors='black')
        ax1.tick_params(axis='y', labelcolor=color, width=1.5, colors=color)
        ax1.spines['left'].set_linewidth(1.5)
        for tick in ax1.xaxis.get_major_ticks():
            tick.label1.set_fontsize(14)
            tick.label1.set_fontfamily('Helvetica')

    if plotbright.get() == 1:
        brightness_s = cycle_data['Brightness_smooth']
        color = '#FF8C00'  # Darker yellow (Orange)
        ax2 = ax1.twinx()
        customize_axis(ax2, color, 'Normalized Greyscale Average')
        ax2.plot(test_time, brightness_s, color=color, linewidth=1.5)
        ax1.spines['right'].set_visible(False)
        axes_list.append(ax2)

    if plotvolt.get() == 1:
        voltage = cycle_data['Voltage(V)_smooth']
        ax3 = ax1.twinx()
        color = 'forestgreen'
        customize_axis(ax3, color, 'Voltage (V)')
        ax3.plot(test_time, voltage, color=color, linewidth=1.5)
        axes_list.append(ax3)

    if plotderiv.get() == 1:
        derivative_s = cycle_data['Brightness Derivative_smooth']
        color = 'firebrick'
        ax4 = ax1.twinx()
        customize_axis(ax4, color, 'NGA Derivative')
        ax4.plot(test_time, derivative_s, color=color, linewidth=1.5)
        axes_list.append(ax4)

    adjust_plot_layout(fig, axes_list)

    # Create a directory "Graphs" within the same directory as the data file
    output_dir = os.path.join(os.path.dirname(filepath), 'Graphs')
    os.makedirs(output_dir, exist_ok=True)

    # Save the figure with a transparent background and unique filename
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_file = os.path.join(output_dir, f'plot_{timestamp}.png')
    fig.savefig(output_file, dpi=200, transparent=True, bbox_inches='tight')

    plt.show()

def select_file():
    filename = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if filename:
        selected_file.set(filename)
        plotcurrent.set(1)
        plotbright.set(1)
        plotvolt.set(0)
        plotderiv.set(0)
        update_cycles(filename)

def update_cycles(filepath):
    data = pd.read_csv(filepath)
    unique_cycles = sorted(data['Cycle_Index'].unique())
    cycle_listbox.delete(0, END)
    for cycle in unique_cycles:
        display_cycle = 'Rest' if cycle == 0 else cycle
        cycle_listbox.insert(END, display_cycle)
    
    # Select all cycles by default
    for i in range(len(unique_cycles)):
        cycle_listbox.selection_set(i)

def create_plot():
    selected_indices = cycle_listbox.curselection()
    cycles = [int(cycle_listbox.get(i)) if cycle_listbox.get(i) != 'Rest' else 0 for i in selected_indices]
    main(selected_file.get(), cycles, plotcurrent, plotbright, plotvolt, plotderiv)

# Add all GUI elements to the main window
Label(root, text="Select a data file to begin:").pack(pady=10)
Button(root, text="Select File", command=select_file).pack(pady=10)
Label(root, textvariable=selected_file).pack(pady=10)

Label(root, text="Select Cycles:").pack()
cycle_listbox = Listbox(root, selectmode=MULTIPLE, exportselection=False)
cycle_listbox.pack(side="left", fill="y", padx=10)

scrollbar = Scrollbar(root, orient="vertical")
scrollbar.config(command=cycle_listbox.yview)
scrollbar.pack(side="left", fill="y")
cycle_listbox.config(yscrollcommand=scrollbar.set)

Label(root, text="Select Data to Plot:").pack(pady=10)
Checkbutton(root, text="Current", variable=plotcurrent).pack()
Checkbutton(root, text="Brightness", variable=plotbright).pack()
Checkbutton(root, text="Voltage", variable=plotvolt).pack()
Checkbutton(root, text="Derivative", variable=plotderiv).pack()

Button(root, text="Create Plot", command=create_plot).pack(pady=10)

root.mainloop()
