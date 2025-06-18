import tkinter as tk
from tkinter import ttk, messagebox, filedialog, Toplevel
import serial
import serial.tools.list_ports
import pandas as pd
import time
from datetime import datetime
# import matplotlib # Remove Matplotlib import
# matplotlib.use('TkAgg') # Remove Matplotlib use
# import matplotlib.pyplot as plt # Remove Matplotlib pyplot
# import matplotlib.animation as animation # Remove Matplotlib animation
# from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk # Remove Matplotlib backends
from collections import deque
import threading
import queue
import os
import shutil

# Import Plotly
import plotly.graph_objects as go
from PIL import Image, ImageTk # For displaying Plotly images in Tkinter
import io # To capture image bytes from Plotly

# --- Heat Index Calculation (remains unchanged) ---
def calculate_heat_index(temp_c, rh):
    # Convert Celsius to Fahrenheit
    temp_f = (temp_c * 9/5) + 32
    rh = max(0, min(100, rh))  # Ensure relative humidity is between 0 and 100

    # Coefficients from the simplified heat index formula (Rothfusz regression)
    c1 = -42.379
    c2 = 2.04901523
    c3 = 10.14333127
    c4 = -0.22475541
    c5 = -6.83783e-3
    c6 = -5.4817e-2
    c7 = 1.22874e-3
    c8 = 8.5282e-4
    c9 = -1.99e-6

    # Heat index calculation
    hi_f = c1 + (c2 * temp_f) + (c3 * rh) + (c4 * temp_f * rh) + \
           (c5 * temp_f**2) + (c6 * rh**2) + (c7 * temp_f**2 * rh) + \
           (c8 * temp_f * rh**2) + (c9 * temp_f**2 * rh**2)

    # Adjustment for lower temperatures and humidity
    if temp_f <= 40:
        hi_f = 0.5 * (temp_f + 61.0 + ((temp_f - 68.0) * 1.2) + (rh * 0.094))
    elif 80 <= temp_f <= 112 and 13 <= rh <= 85:
        # Use the full regression equation without adjustments
        pass
    elif temp_f > 112 or rh < 13 or rh > 85:
        # Simplified adjustment for extreme conditions
        hi_f = temp_f + (0.25 * (rh - 50)) + (0.5 * (temp_f - 75))

    # Convert back to Celsius
    hi_c = (hi_f - 32) * 5/9
    return round(hi_c, 2)

# --- Configuration ---
columns = ['Device Timestamp', 'Temperature (°C)', 'Humidity (%)', 'Temp2 (°C)', 'Humidity2 (%)', 'Heat Index (°C)', 'Heat Index2 (°C)']
new_records_buffer = []
dialog_checkboxes = {}  # Track dialog checkboxes

COM_PORT = ''
BAUD_RATE = 9600
SAVE_INTERVAL_RECORDS = 60
SAVE_INTERVAL_SECONDS = 300
PLOT_MAX_POINTS = 300
PLOT_UPDATE_INTERVAL_MS = 1000
SAVE_DIR = os.path.expanduser("~/Desktop")  # Default save directory

plot_timestamps = deque(maxlen=PLOT_MAX_POINTS)
plot_temperatures = deque(maxlen=PLOT_MAX_POINTS)
plot_humidities = deque(maxlen=PLOT_MAX_POINTS)
plot_temp2 = deque(maxlen=PLOT_MAX_POINTS)
plot_humidity2 = deque(maxlen=PLOT_MAX_POINTS)
plot_heat_index = deque(maxlen=PLOT_MAX_POINTS)
plot_heat_index2 = deque(maxlen=PLOT_MAX_POINTS)

last_save_time = time.time()
active_graph = 'temperature'  # Default graph
latest_values = {'temp': 0, 'rh': 0, 'temp2': 0, 'humidity2': 0, 'heat_index': 0, 'heat_index2': 0}
# pan_mode = False # Plotly handles panning internally
# panning = False  # Track if panning is active

# --- Threading Setup ---
stop_event = threading.Event()
data_queue = queue.Queue()
ser = None
# ani = None # Remove Matplotlib animation object

# --- GUI Setup ---
root = tk.Tk()
root.title("Fluke 1620A Data Logger")
root.geometry("1280x720")

# Real-Time Values Frame (Top)
real_time_frame = ttk.Frame(root)
real_time_frame.pack(fill="x", padx=10, pady=5, anchor='n')

# Labels for real-time values
temp1_label = ttk.Label(real_time_frame, text="Temperature-1: ", font=("Arial", 10))
temp1_value_label = ttk.Label(real_time_frame, text="0.00 °C", font=("Arial", 10))
separator1 = ttk.Label(real_time_frame, text=" | ", font=("Arial", 10))
heat_index1_label = ttk.Label(real_time_frame, text="Heat Index-1: ", font=("Arial", 10))
heat_index1_value_label = ttk.Label(real_time_frame, text="0.00 °C", font=("Arial", 10))
separator2 = ttk.Label(real_time_frame, text=" | ", font=("Arial", 10))
temp2_label = ttk.Label(real_time_frame, text="Temperature-2: ", font=("Arial", 10))
temp2_value_label = ttk.Label(real_time_frame, text="0.00 °C", font=("Arial", 10))
separator3 = ttk.Label(real_time_frame, text=" | ", font=("Arial", 10))
heat_index2_label = ttk.Label(real_time_frame, text="Heat Index-2: ", font=("Arial", 10))
heat_index2_value_label = ttk.Label(real_time_frame, text="0.00 °C", font=("Arial", 10))

# Pack labels horizontally
temp1_label.pack(side="left")
temp1_value_label.pack(side="left")
separator1.pack(side="left")
heat_index1_label.pack(side="left")
heat_index1_value_label.pack(side="left")
separator2.pack(side="left")
temp2_label.pack(side="left")
temp2_value_label.pack(side="left")
separator3.pack(side="left")
heat_index2_label.pack(side="left")
heat_index2_value_label.pack(side="left")

# Parameter Frame
param_frame = ttk.Frame(root, padding="10")
param_frame.pack(fill="x")

# COM Port Dropdown
ttk.Label(param_frame, text="COM Port:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
try:
    com_ports = [port.device for port in serial.tools.list_ports.comports()]
except Exception as e:
    com_ports = []
    messagebox.showerror("Error", f"Failed to list COM ports: {e}")
com_port_var = tk.StringVar()
com_port_combo = ttk.Combobox(param_frame, textvariable=com_port_var, values=com_ports, state="readonly", width=15)
com_port_combo.grid(row=0, column=1, padx=5, pady=5)
if com_ports:
    com_port_combo.current(0)
else:
    com_port_var.set("No ports available")

# Other Parameters
ttk.Label(param_frame, text="Baud Rate:").grid(row=0, column=2, sticky="e", padx=15, pady=5)
baud_rate_var = tk.StringVar(value=str(BAUD_RATE))
baud_rate_entry = ttk.Entry(param_frame, textvariable=baud_rate_var, width=15)
baud_rate_entry.grid(row=0, column=3, padx=5, pady=5)

ttk.Label(param_frame, text="Save Interval Records:").grid(row=0, column=4, sticky="e", padx=15, pady=5)
save_interval_records_var = tk.StringVar(value=str(SAVE_INTERVAL_RECORDS))
save_interval_records_entry = ttk.Entry(param_frame, textvariable=save_interval_records_var, width=15)
save_interval_records_entry.grid(row=0, column=5, padx=5, pady=5)

ttk.Label(param_frame, text="Save Interval Seconds:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
save_interval_seconds_var = tk.StringVar(value=str(SAVE_INTERVAL_SECONDS))
save_interval_seconds_entry = ttk.Entry(param_frame, textvariable=save_interval_seconds_var, width=15)
save_interval_seconds_entry.grid(row=1, column=1, padx=5, pady=5)

ttk.Label(param_frame, text="Plot Max Points:").grid(row=1, column=2, sticky="e", padx=15, pady=5)
plot_max_points_var = tk.StringVar(value=str(PLOT_MAX_POINTS))
plot_max_points_entry = ttk.Entry(param_frame, textvariable=plot_max_points_var, width=15)
plot_max_points_entry.grid(row=1, column=3, padx=5, pady=5)

ttk.Label(param_frame, text="Plot Update Interval (ms):").grid(row=1, column=4, sticky="e", padx=15, pady=5)
plot_update_interval_ms_var = tk.StringVar(value=str(PLOT_UPDATE_INTERVAL_MS))
plot_update_interval_ms_entry = ttk.Entry(param_frame, textvariable=plot_update_interval_ms_var, width=15)
plot_update_interval_ms_entry.grid(row=1, column=5, padx=5, pady=5)

# Save Directory Selection
ttk.Label(param_frame, text="Save Directory:").grid(row=2, column=2, sticky="e", padx=5, pady=5)
save_dir_var = tk.StringVar(value=SAVE_DIR)
ttk.Entry(param_frame, textvariable=save_dir_var, width=15, state="readonly").grid(row=2, column=3, padx=5, pady=5)
ttk.Button(param_frame, text="Browse", command=lambda: browse_directory(save_dir_var)).grid(row=2, column=4, padx=5, pady=5)

# Button Frame
button_frame = ttk.Frame(root, padding="10")
button_frame.pack(fill="x")

start_button = ttk.Button(button_frame, text="Start Logging", command=lambda: start_logging())
start_button.pack(side="left", padx=10, pady=5)

stop_button = ttk.Button(button_frame, text="Stop Logging", command=lambda: stop_logging(), state="disabled")
stop_button.pack(side="left", padx=10, pady=5)

# pan_button = ttk.Button(button_frame, text="Pan Graph", command=lambda: toggle_pan_mode()) # Remove pan button (Plotly handles it)
# pan_button.pack(side="left", padx=10, pady=5)

# Graph Toggle Buttons with Checkboxes
toggle_frame = ttk.Frame(root, padding="10")
toggle_frame.pack(fill="x")

graphs = ['temperature', 'humidity', 'temp2', 'humidity2', 'heat_index', 'heat_index2']
dialog_vars = {graph: tk.BooleanVar(value=False) for graph in graphs}

def set_graph(graph):
    global active_graph
    active_graph = graph
    update_main_graph() # Call update for the main graph

for graph in graphs:
    frame = ttk.Frame(toggle_frame)
    frame.pack(side="left", padx=5)
    ttk.Button(frame, text=graph.capitalize(), command=lambda g=graph: set_graph(g)).pack(side="top")
    dialog_checkboxes[graph] = ttk.Checkbutton(frame, text="Open in dialog", variable=dialog_vars[graph], command=lambda g=graph: toggle_dialog(g))
    dialog_checkboxes[graph].pack(side="bottom")

# Status Label
status_var = tk.StringVar(value="Ready")
status_label = ttk.Label(root, textvariable=status_var, padding="5")
status_label.pack(fill="x")

# Coordinate Label (Bottom Left) - Not directly supported by static Plotly embedding
# coord_frame = ttk.Frame(root)
# coord_frame.pack(anchor="sw", padx=10, pady=5)
# coord_var = tk.StringVar(value="X: -, Y: -")
# coord_label = ttk.Label(coord_frame, textvariable=coord_var, font=("Arial", 10))
# coord_label.pack()

# Plot Frame
plot_frame = ttk.Frame(root)
plot_frame.pack(fill="both", expand=True)

# Plotly Figure and Tkinter Canvas for main plot
main_fig = go.Figure()
main_fig.update_layout(
    title='Real-Time Data',
    xaxis_title='Time',
    yaxis_title='Temperature (°C)',
    # hovermode='x unified' # Enable hover for interactive features if a web view was used
)
# Add an initial trace
main_fig.add_trace(go.Scatter(x=[], y=[], mode='lines', name='Temperature (°C)', line=dict(color='red')))
main_graph_image_label = tk.Label(plot_frame)
main_graph_image_label.pack(fill="both", expand=True)

# Removed Matplotlib specific toolbar and pan functions
# toolbar = NavigationToolbar2Tk(canvas, plot_frame)
# toolbar.update()
# def toggle_pan_mode(): # Not needed with Plotly's built-in interactivity
#     global pan_mode
#     pan_mode = not pan_mode
#     if pan_mode:
#         canvas.toolbar.pan()
#         pan_button.config(text="Stop Panning")
#     else:
#         canvas.toolbar.pan()
#         canvas.toolbar.home()
#         pan_button.config(text="Pan Graph")
#     canvas.draw()
# canvas.mpl_connect('motion_notify_event', on_mouse_move) # Not needed

# Dialog windows for graphs
dialog_windows = {}

def toggle_dialog(graph):
    if dialog_vars[graph].get():
        if graph not in dialog_windows:
            dialog_windows[graph] = Toplevel(root)
            dialog_windows[graph].title(f"{graph.capitalize()} Graph")
            dialog_windows[graph].geometry("800x600")
            dialog_windows[graph].protocol("WM_DELETE_WINDOW", lambda g=graph: close_dialog(g))

            fig_dialog = go.Figure()
            fig_dialog.update_layout(
                title=f"{graph.capitalize()} Graph",
                xaxis_title='Time',
                yaxis_title=get_label(graph),
                # hovermode='x unified'
            )
            fig_dialog.add_trace(go.Scatter(x=[], y=[], mode='lines', name=graph.capitalize(), line=dict(color=get_color(graph))))

            dialog_windows[graph].plotly_figure = fig_dialog
            dialog_windows[graph].image_label = tk.Label(dialog_windows[graph])
            dialog_windows[graph].image_label.pack(fill="both", expand=True)

            # Schedule the update for the dialog graph
            dialog_windows[graph].after_id = dialog_windows[graph].after(PLOT_UPDATE_INTERVAL_MS, lambda g=graph: update_dialog_graph(g))
    else:
        close_dialog(graph)

def close_dialog(graph):
    if graph in dialog_windows:
        # Cancel the scheduled update
        if hasattr(dialog_windows[graph], 'after_id'):
            dialog_windows[graph].after_cancel(dialog_windows[graph].after_id)
        dialog_windows[graph].destroy()
        del dialog_windows[graph]
        dialog_vars[graph].set(False)

# Helper functions
def get_label(graph):
    labels = {
        'temperature': 'Temperature (°C)',
        'humidity': 'Humidity (%)',
        'temp2': 'Temp2 (°C)',
        'humidity2': 'Humidity2 (%)',
        'heat_index': 'Heat Index (°C)',
        'heat_index2': 'Heat Index2 (°C)'
    }
    return labels[graph]

def get_ylim(graph):
    ylims = {
        'temperature': (20, 30),
        'humidity': (40, 60),
        'temp2': (-10, 30),
        'humidity2': (0, 100),
        'heat_index': (20, 40),
        'heat_index2': (0, 40)
    }
    return ylims[graph]

def get_color(graph):
    colors = {
        'temperature': 'red',
        'humidity': 'blue',
        'temp2': 'green',
        'humidity2': 'magenta',
        'heat_index': "orange",
        'heat_index2': 'purple'
    }
    return colors[graph]

# --- Functions ---
def update_main_graph():
    # Update main_fig layout
    main_fig.update_layout(
        yaxis_title=get_label(active_graph),
        yaxis_range=get_ylim(active_graph),
        showlegend=True # Ensure legend is shown
    )
    # Update the trace data and color
    main_fig.data[0].name = get_label(active_graph)
    main_fig.data[0].line.color = get_color(active_graph)

    # Re-draw the graph
    update_plot_image(main_fig, main_graph_image_label)

def browse_directory(save_dir_var):
    global SAVE_DIR
    new_dir = filedialog.askdirectory(initialdir=SAVE_DIR, title="Select Save Directory")
    if new_dir:
        SAVE_DIR = new_dir
        save_dir_var.set(SAVE_DIR)

def start_logging():
    global ser, new_records_buffer, plot_timestamps, plot_temperatures, plot_humidities, plot_temp2, plot_humidity2
    global plot_heat_index, plot_heat_index2, last_save_time, stop_event, data_queue, COM_PORT, BAUD_RATE
    global SAVE_INTERVAL_RECORDS, SAVE_INTERVAL_SECONDS, PLOT_MAX_POINTS, PLOT_UPDATE_INTERVAL_MS

    # Validate parameters
    COM_PORT = com_port_var.get()
    if not COM_PORT or COM_PORT == "No ports available":
        messagebox.showerror("Input Error", "Please select a valid COM port.")
        return
    try:
        BAUD_RATE = int(baud_rate_var.get())
        if BAUD_RATE <= 0:
            raise ValueError("Baud Rate must be positive")
        SAVE_INTERVAL_RECORDS = int(save_interval_records_var.get())
        if SAVE_INTERVAL_RECORDS <= 0:
            raise ValueError("Save Interval Records must be positive")
        SAVE_INTERVAL_SECONDS = int(save_interval_seconds_var.get())
        if SAVE_INTERVAL_SECONDS <= 0:
            raise ValueError("Save Interval Seconds must be positive")
        PLOT_MAX_POINTS = int(plot_max_points_var.get())
        if PLOT_MAX_POINTS <= 0:
            raise ValueError("Plot Max Points must be positive")
        PLOT_UPDATE_INTERVAL_MS = int(plot_update_interval_ms_var.get())
        if PLOT_UPDATE_INTERVAL_MS <= 0:
            raise ValueError("Plot Update Interval must be positive")
    except ValueError as e:
        messagebox.showerror("Input Error", f"Invalid input: {e}")
        return

    # Reset data structures
    new_records_buffer.clear()
    plot_timestamps = deque(maxlen=PLOT_MAX_POINTS)
    plot_temperatures = deque(maxlen=PLOT_MAX_POINTS)
    plot_humidities = deque(maxlen=PLOT_MAX_POINTS)
    plot_temp2 = deque(maxlen=PLOT_MAX_POINTS)
    plot_humidity2 = deque(maxlen=PLOT_MAX_POINTS)
    plot_heat_index = deque(maxlen=PLOT_MAX_POINTS)
    plot_heat_index2 = deque(maxlen=PLOT_MAX_POINTS)
    last_save_time = time.time()
    stop_event.clear()
    data_queue = queue.Queue()
    latest_values.update({'temp': 0, 'rh': 0, 'temp2': 0, 'humidity2': 0, 'heat_index': 0, 'heat_index2': 0})

    # Start serial thread
    status_var.set("Starting serial connection...")
    serial_thread = threading.Thread(target=serial_reader_thread, daemon=True)
    serial_thread.start()

    # Schedule the Plotly update function
    root.after(PLOT_UPDATE_INTERVAL_MS, update_plots) # Start the plot update loop

    # Update UI
    start_button.config(state="disabled")
    stop_button.config(state="normal")
    status_var.set("Logging started")
    messagebox.showinfo("Info", "Logging started.")

def stop_logging():
    global ser, stop_event
    stop_event.set()
    # Cancel the scheduled Plotly update
    root.after_cancel(root.after_id) # Assuming root.after_id is set when update_plots is scheduled

    if ser and hasattr(ser, 'is_open') and ser.is_open:
        try:
            ser.close()
            status_var.set("Serial port closed")
        except Exception as e:
            status_var.set(f"Error closing serial port: {e}")
    else:
        status_var.set("No active serial connection to close")

    # Always save remaining data
    if new_records_buffer:
        status_var.set(f"Saving {len(new_records_buffer)} remaining records...")
        save_to_excel(new_records_buffer)
        new_records_buffer.clear()
    else:
        status_var.set("No new records to save")
        messagebox.showinfo("Info", "No new records to save")

    # Update UI
    start_button.config(state="normal")
    stop_button.config(state="disabled")

def on_closing():
    if messagebox.askokcancel("Quit", "Do you want to quit? Logging will stop."):
        stop_logging()
        root.destroy()

root.protocol("WM_DELETE_WINDOW", on_closing)

def serial_reader_thread():
    global ser, last_save_time
    try:
        ser = serial.Serial(COM_PORT, BAUD_RATE, timeout=1)
        status_var.set(f"Connected to {COM_PORT} at {BAUD_RATE} baud")
    except serial.SerialException as se:
        status_var.set(f"Serial connection failed: {se}")
        messagebox.showerror("Serial Error", f"Failed to connect to {COM_PORT}: {se}")
        stop_event.set()
        return

    while not stop_event.is_set():
        try:
            line = ser.readline().decode(errors='ignore').strip()
            if not line:
                continue

            cleaned_line = line.replace('\xa0', ' ').strip()
            parts = [part.strip() for part in cleaned_line.split(',')]

            if len(parts) >= 8:
                device_time_str = parts[0]
                try:
                    temp = float(parts[1])
                    rh = float(parts[3])
                    temp2 = float(parts[5])
                    humidity2 = float(parts[7])

                    heat_index = calculate_heat_index(temp, rh)
                    heat_index2 = calculate_heat_index(temp2, humidity2)

                    data_queue.put({
                        'timestamp_str': device_time_str,
                        'temp': temp,
                        'rh': rh,
                        'temp2': temp2,
                        'humidity2': humidity2,
                        'heat_index': heat_index,
                        'heat_index2': heat_index2
                    })

                    status_var.set(f"Logged: {device_time_str}, T={temp}°C, RH={rh}%")

                except (ValueError, IndexError) as e:
                    status_var.set(f"Data parsing error: {e}")

        except serial.SerialException as se:
            status_var.set(f"Serial error: {se}. Attempting reconnect...")
            try:
                if ser and ser.is_open:
                    ser.close()
                ser = serial.Serial(COM_PORT, BAUD_RATE, timeout=1)
                status_var.set("Reconnected")
            except Exception as re_e:
                status_var.set(f"Reconnect failed: {re_e}")
                time.sleep(1)

        except Exception as e:
            status_var.set(f"Unexpected error: {e}")

    if ser and ser.is_open:
        try:
            ser.close()
            status_var.set("Serial port closed")
        except Exception as e:
            status_var.set(f"Error closing serial port: {e}")

# This function replaces the Matplotlib animate function
def update_plots():
    global last_save_time, new_records_buffer, latest_values

    # Process data from queue
    while not data_queue.empty():
        data = data_queue.get()
        try:
            dt_object = datetime.strptime(data['timestamp_str'], '%d/%m/%Y %H:%M:%S')
        except ValueError as ve:
            status_var.set(f"Timestamp parse error: {ve}")
            # Re-schedule the update even if there's an error, to keep the loop going
            root.after_id = root.after(PLOT_UPDATE_INTERVAL_MS, update_plots)
            return

        plot_timestamps.append(dt_object)
        plot_temperatures.append(data['temp'])
        plot_humidities.append(data['rh'])
        plot_temp2.append(data['temp2'])
        plot_humidity2.append(data['humidity2'])
        plot_heat_index.append(data['heat_index'])
        plot_heat_index2.append(data['heat_index2'])

        latest_values.update({
            'temp': data['temp'],
            'rh': data['rh'],
            'temp2': data['temp2'],
            'humidity2': data['humidity2'],
            'heat_index': data['heat_index'],
            'heat_index2': data['heat_index2']
        })

        # Update real-time values
        update_real_time_values()

        new_records_buffer.append([
            data['timestamp_str'], data['temp'], data['rh'],
            data['temp2'], data['humidity2'], data['heat_index'], data['heat_index2']
        ])

    # Update main plot
    data_map = {
        'temperature': plot_temperatures,
        'humidity': plot_humidities,
        'temp2': plot_temp2,
        'humidity2': plot_humidity2,
        'heat_index': plot_heat_index,
        'heat_index2': plot_heat_index2
    }
    if plot_timestamps:
        # Update the main Plotly figure's trace data
        main_fig.data[0].x = list(plot_timestamps)
        main_fig.data[0].y = list(data_map[active_graph])

        # Convert Plotly figure to image and update Tkinter label
        update_plot_image(main_fig, main_graph_image_label)

    # Update dialog plots
    for graph, dialog_win in dialog_windows.items():
        if dialog_win.winfo_exists(): # Check if dialog window is still open
            dialog_fig = dialog_win.plotly_figure
            dialog_fig.data[0].x = list(plot_timestamps)
            dialog_fig.data[0].y = list(data_map[graph])
            update_plot_image(dialog_fig, dialog_win.image_label)
        else:
            # If window was closed externally, clean up
            close_dialog(graph)


    # Save periodically
    current_time = time.time()
    if (len(new_records_buffer) >= SAVE_INTERVAL_RECORDS) or \
       (new_records_buffer and (current_time - last_save_time) >= SAVE_INTERVAL_SECONDS):
        status_var.set(f"Saving {len(new_records_buffer)} records...")
        save_to_excel(new_records_buffer)
        new_records_buffer.clear()
        last_save_time = current_time

    # Re-schedule the update
    root.after_id = root.after(PLOT_UPDATE_INTERVAL_MS, update_plots)

def update_plot_image(plotly_fig, tk_image_label):
    # Get the current size of the Tkinter label to render the Plotly image with appropriate dimensions
    width = tk_image_label.winfo_width()
    height = tk_image_label.winfo_height()

    if width == 0 or height == 0: # If widget is not yet rendered, use a default size
        width, height = 800, 400

    try:
        img_bytes = plotly_fig.to_image(format="png", width=width, height=height)
        pil_image = Image.open(io.BytesIO(img_bytes))
        tk_image = ImageTk.PhotoImage(pil_image)
        tk_image_label.config(image=tk_image)
        tk_image_label.image = tk_image  # Keep a reference!
    except Exception as e:
        status_var.set(f"Error rendering plot: {e}")

def update_dialog_graph(graph):
    # This function is now mostly handled by update_plots, but the after schedule needs to be called
    # to keep the dialog's update loop active if it's separate from the main root loop.
    # We call update_plots which iterates through all dialogs.
    if graph in dialog_windows and dialog_windows[graph].winfo_exists():
        dialog_windows[graph].after_id = dialog_windows[graph].after(PLOT_UPDATE_INTERVAL_MS, lambda g=graph: update_dialog_graph(g))

def update_real_time_values():
    # Update Temperature-1
    temp1 = latest_values['temp']
    temp1_value_label.config(text=f"{temp1:.2f} °C")

    # Update Heat Index-1
    heat_index1 = latest_values['heat_index']
    heat_index1_value_label.config(text=f"{heat_index1:.2f} °C")
    if abs(heat_index1 - temp1) <= 5:
        heat_index1_value_label.config(foreground="green")
    else:
        heat_index1_value_label.config(foreground="red")

    # Update Temperature-2
    temp2 = latest_values['temp2']
    temp2_value_label.config(text=f"{temp2:.2f} °C")

    # Update Heat Index-2
    heat_index2 = latest_values['heat_index2']
    heat_index2_value_label.config(text=f"{heat_index2:.2f} °C")
    if abs(heat_index2 - temp2) <= 5:
        heat_index2_value_label.config(foreground="green")
    else:
        heat_index2_value_label.config(foreground="red")

def save_to_excel(records):
    date_str = datetime.now().strftime("%Y%m%d")
    excel_file = os.path.join(SAVE_DIR, f"fluke_1620A_{date_str}.xlsx")
    temp_file = os.path.join(SAVE_DIR, f"fluke_1620A_{date_str}_tmp.xlsx")

    if not records:
        return

    temp_df = pd.DataFrame(records, columns=columns)
    try:
        if os.path.exists(excel_file):
            try:
                existing_df = pd.read_excel(excel_file, engine='openpyxl')
                updated_df = pd.concat([existing_df, temp_df], ignore_index=True)
            except Exception as e:
                status_var.set(f"Error reading existing file: {e}. Creating new file.")
                updated_df = temp_df
        else:
            updated_df = temp_df

        updated_df.to_excel(temp_file, index=False, engine='openpyxl')
        if os.path.exists(temp_file):
            if os.path.exists(excel_file):
                os.remove(excel_file)
            shutil.move(temp_file, excel_file)
        status_var.set(f"Saved to {excel_file}. Total records: {len(updated_df)}")
    except Exception as e:
        status_var.set(f"Save failed: {e}")
        messagebox.showerror("Error", f"Excel save error: {e}")
    finally:
        if os.path.exists(temp_file):
            os.remove(temp_file)  # Clean up temp file

# --- Start GUI ---
root.mainloop()
