import tkinter as tk
from tkinter import ttk, messagebox, filedialog, Toplevel
import serial
import serial.tools.list_ports
import pandas as pd
import time
from datetime import datetime
import matplotlib
matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
import matplotlib.animation as animation
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from collections import deque
import threading
import queue
import os
import shutil

# --- Heat Index Calculation ---
def calculate_heat_index(temp_c, rh):
    temp_f = (temp_c * 9/5) + 32
    rh = max(0, min(100, rh))
    c1 = -42.379
    c2 = 2.04901523
    c3 = 10.14333127
    c4 = -0.22475541
    c5 = -6.83783e-3
    c6 = -5.4817e-2
    c7 = 1.22874e-3
    c8 = 8.5282e-4
    c9 = -1.99e-6
    hi_f = c1 + (c2 * temp_f) + (c3 * rh) + (c4 * temp_f * rh) + \
           (c5 * temp_f**2) + (c6 * rh**2) + (c7 * temp_f**2 * rh) + \
           (c8 * temp_f * rh**2) + (c9 * temp_f**2 * rh**2)
    if temp_f <= 40:
        hi_f = 0.5 * (temp_f + 61.0 + ((temp_f - 68.0) * 1.2) + (rh * 0.094))
    elif 80 <= temp_f <= 112 and 13 <= rh <= 85:
        pass
    elif temp_f > 112 or rh < 13 or rh > 85:
        hi_f = temp_f + (0.25 * (rh - 50)) + (0.5 * (temp_f - 75))
    hi_c = (hi_f - 32) * 5/9
    return round(hi_c, 2)

# --- Configuration ---
columns = ['Device Timestamp', 'Temperature (°C)', 'Humidity (%)', 'Temp2 (°C)', 'Humidity2 (%)', 'Heat Index (°C)', 'Heat Index2 (°C)']
new_records_buffer = []
dialog_checkboxes = {}

COM_PORT = ''
BAUD_RATE = 9600
SAVE_INTERVAL_RECORDS = 60
SAVE_INTERVAL_SECONDS = 300
PLOT_MAX_POINTS = 300
PLOT_UPDATE_INTERVAL_MS = 100
SAVE_DIR = os.path.expanduser("~/Desktop")

plot_timestamps = deque(maxlen=PLOT_MAX_POINTS)
plot_temperatures = deque(maxlen=PLOT_MAX_POINTS)
plot_humidities = deque(maxlen=PLOT_MAX_POINTS)
plot_temp2 = deque(maxlen=PLOT_MAX_POINTS)
plot_humidity2 = deque(maxlen=PLOT_MAX_POINTS)
plot_heat_index = deque(maxlen=PLOT_MAX_POINTS)
plot_heat_index2 = deque(maxlen=PLOT_MAX_POINTS)

last_save_time = time.time()
active_graph = 'temperature'
latest_values = {'temp': 0, 'rh': 0, 'temp2': 0, 'humidity2': 0, 'heat_index': 0, 'heat_index2': 0}
pan_mode = False
panning = False

# --- Threading Setup ---
stop_event = threading.Event()
data_queue = queue.Queue()
ser = None
ani = None

# --- GUI Setup ---
root = tk.Tk()
root.title("Fluke 1620A Data Logger")
root.geometry("1280x720")
root.configure(bg="#ffffff")

style = ttk.Style()
style.theme_use('clam')
style.configure("TLabel", background="#ffffff", font=("Arial", 12))
style.configure("TButton", padding=6, font=("Arial", 10), background="#ffffff")
style.configure("TCombobox", padding=5, font=("Arial", 10), background="#ffffff")
style.configure("Card.TFrame", background="#ffffff", relief="groove", borderwidth=2)
style.configure("TLabelframe", background="#ffffff")
style.configure("TLabelframe.Label", background="#ffffff")

# Banner
banner_frame = ttk.Frame(root, style='TLabel')
banner_frame.pack(fill="x", pady=(10, 5))
banner_label = ttk.Label(banner_frame, text="NPL Data Logger", font=("Arial", 20, "bold"), foreground="#2c3e50")
banner_label.pack()
ttk.Separator(root, orient="horizontal").pack(fill="x", pady=5)

# Main Frame with Bordered Panels
main_frame = ttk.Frame(root, padding=10, style="Card.TFrame")
main_frame.pack(fill="both", expand=True)
left_panel = ttk.Frame(main_frame, padding=10, style="Card.TFrame")
left_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
right_panel = ttk.Frame(main_frame, padding=10, style="Card.TFrame")
right_panel.grid(row=0, column=1, sticky="nsew")
main_frame.columnconfigure(1, weight=3)

# --- Enhanced Real-Time Values Section ---
real_time_frame = ttk.LabelFrame(left_panel, text="Real-Time Values", padding="10")
real_time_frame.pack(fill="x", pady=(0, 20))

highlight_style = ("Arial", 32, "bold")
label_style = ("Arial", 10, "bold")

value_labels = {
    "Temperature-1": (ttk.Label(real_time_frame, text="Temperature-1", font=label_style),
                      ttk.Label(real_time_frame, text="0.00 °C", font=highlight_style, foreground="#e74c3c")),
    "Heat Index-1": (ttk.Label(real_time_frame, text="Heat Index-1", font=label_style),
                     ttk.Label(real_time_frame, text="0.00 °C", font=highlight_style, foreground="#2ecc71")),
    "Temperature-2": (ttk.Label(real_time_frame, text="Temperature-2", font=label_style),
                      ttk.Label(real_time_frame, text="0.00 °C", font=highlight_style, foreground="#3498db")),
    "Heat Index-2": (ttk.Label(real_time_frame, text="Heat Index-2", font=label_style),
                     ttk.Label(real_time_frame, text="0.00 °C", font=highlight_style, foreground="#9b59b6"))
}

for i, (label, (lbl, val)) in enumerate(value_labels.items()):
    lbl.grid(row=i, column=0, sticky="e", pady=10, padx=5)
    val.grid(row=i, column=1, sticky="w", pady=10, padx=5)

(temp1_label, temp1_value_label) = value_labels["Temperature-1"]
(heat_index1_label, heat_index1_value_label) = value_labels["Heat Index-1"]
(temp2_label, temp2_value_label) = value_labels["Temperature-2"]
(heat_index2_label, heat_index2_value_label) = value_labels["Heat Index-2"]

# Parameter Frame
param_frame = ttk.Frame(left_panel, padding="10", style="Card.TFrame")
param_frame.pack(fill="x")

# COM Port Dropdown
ttk.Label(param_frame, text="COM Port:", style="TLabel").grid(row=0, column=0, sticky="e", padx=5, pady=5)
try:
    com_ports = [port.device for port in serial.tools.list_ports.comports()]
except Exception as e:
    com_ports = []
    messagebox.showerror("Error", f"Failed to list COM ports: {e}")
com_port_var = tk.StringVar()
com_port_combo = ttk.Combobox(param_frame, textvariable=com_port_var, values=com_ports, state="readonly", width=15, style="TCombobox")
com_port_combo.grid(row=0, column=1, padx=5, pady=5)
if com_ports:
    com_port_combo.current(0)
else:
    com_port_var.set("No ports available")

# Other Parameters
ttk.Label(param_frame, text="Baud Rate:",background="#ffffff").grid(row=1, column=0, sticky="e", padx=5, pady=5)
baud_rate_var = tk.StringVar(value=str(BAUD_RATE))
baud_rate_entry = ttk.Entry(param_frame, textvariable=baud_rate_var, width=15)
baud_rate_entry.grid(row=1, column=1, padx=5, pady=5)

ttk.Label(param_frame, text="Save Interval Records:", background="#ffffff").grid(row=2, column=0, sticky="e", padx=5, pady=5)
save_interval_records_var = tk.StringVar(value=str(SAVE_INTERVAL_RECORDS))
save_interval_records_entry = ttk.Entry(param_frame, textvariable=save_interval_records_var, width=15)
save_interval_records_entry.grid(row=2, column=1, padx=5, pady=5)

ttk.Label(param_frame, text="Save Interval Seconds:", background="#ffffff").grid(row=3, column=0, sticky="e", padx=5, pady=5)
save_interval_seconds_var = tk.StringVar(value=str(SAVE_INTERVAL_SECONDS))
save_interval_seconds_entry = ttk.Entry(param_frame, textvariable=save_interval_seconds_var, width=15)
save_interval_seconds_entry.grid(row=3, column=1, padx=5, pady=5)

ttk.Label(param_frame, text="Plot Max Points:", background="#ffffff").grid(row=4, column=0, sticky="e", padx=5, pady=5)
plot_max_points_var = tk.StringVar(value=str(PLOT_MAX_POINTS))
plot_max_points_entry = ttk.Entry(param_frame, textvariable=plot_max_points_var, width=15)
plot_max_points_entry.grid(row=4, column=1, padx=5, pady=5)

ttk.Label(param_frame, text="Plot Update Interval (ms):", background="#ffffff").grid(row=5, column=0, sticky="e", padx=15, pady=5)
plot_update_interval_ms_var = tk.StringVar(value=str(PLOT_UPDATE_INTERVAL_MS))
plot_update_interval_ms_entry = ttk.Entry(param_frame, textvariable=plot_update_interval_ms_var, width=15)
plot_update_interval_ms_entry.grid(row=5, column=1, padx=5, pady=5)

# Save Directory Selection
ttk.Label(param_frame, text="Save Directory:", background="#ffffff").grid(row=6, column=0, sticky="e", padx=5, pady=5)
save_dir_var = tk.StringVar(value=SAVE_DIR)
ttk.Entry(param_frame, textvariable=save_dir_var, width=15, state="readonly").grid(row=6, column=1, padx=5, pady=5)
ttk.Button(param_frame, text="Browse", command=lambda: browse_directory(save_dir_var)).grid(row=6, column=2, padx=5, pady=5)

# Button Frame
button_frame = ttk.Frame(root, padding="10")
button_frame.pack(fill="x")

start_button = ttk.Button(button_frame, text="Start Logging", command=lambda: start_logging())
start_button.pack(side="left", padx=10, pady=5)

stop_button = ttk.Button(button_frame, text="Stop Logging", command=lambda: stop_logging(), state="disabled")
stop_button.pack(side="left", padx=10, pady=5)

calibrate_button = ttk.Button(button_frame, text="Calibrate Time", command=lambda: calibrate_time())
calibrate_button.pack(side="left", padx=10, pady=5)

# Graph Toggles
toggle_frame = ttk.LabelFrame(left_panel, text="Graph Selection", padding="5")
toggle_frame.pack(fill="x")
graphs = ['temperature', 'humidity', 'temp2', 'humidity2', 'heat_index', 'heat_index2']
dialog_vars = {graph: tk.BooleanVar(value=False) for graph in graphs}
for i, graph in enumerate(graphs):
    frame = ttk.Frame(toggle_frame)
    frame.grid(row=i//3, column=i%3, padx=5, pady=2)
    ttk.Button(frame, text=graph.capitalize(), command=lambda g=graph: set_graph(g)).pack(side="top")
    dialog_checkboxes[graph] = ttk.Checkbutton(frame, text="Dialog", variable=dialog_vars[graph], command=lambda g=graph: toggle_dialog(g))
    dialog_checkboxes[graph].pack(side="bottom")

# Right Panel (Plot)
right_panel = ttk.Frame(main_frame, padding="10")
right_panel.grid(row=0, column=1, sticky="nsew")
main_frame.columnconfigure(1, weight=3)

fig, ax = plt.subplots(figsize=(10, 6))
fig.suptitle('Real-Time Data', fontsize=14)
ax.set_xlabel('Time')
ax.set_ylabel('Temperature (°C)')
ax.grid(True)
line, = ax.plot([], [], 'r-', label='Temperature (°C)')
ax.legend(loc='upper left')
ax.set_ylim(20, 30)
fig.autofmt_xdate()
fig.tight_layout()

canvas = FigureCanvasTkAgg(fig, master=right_panel)
canvas.get_tk_widget().pack(fill="both", expand=True)
toolbar = NavigationToolbar2Tk(canvas, right_panel)
toolbar.pack()

# Status and Bottom Frame
status_frame = ttk.Frame(root, padding="5")
status_frame.pack(fill="x", side="bottom")
status_var = tk.StringVar(value="Ready")
ttk.Label(status_frame, textvariable=status_var).pack(side="left", padx=10)
coord_var = tk.StringVar(value="X: -, Y: -")
clock_var = tk.StringVar(value=datetime.now().strftime("%H:%M:%S"))
ttk.Label(banner_frame, textvariable=clock_var, font=("Arial", 30), foreground="#00821c").pack(anchor='center', padx=10, pady= 10)

def update_clock():
    clock_var.set(datetime.now().strftime("%H:%M:%S"))
    root.after(1000, update_clock)
update_clock()


def on_mouse_move(event):
    if event.inaxes and not panning:
        x, y = event.xdata, event.ydata
        coord_var.set(f"X: {x:.2f}, Y: {y:.2f}" if x is not None and y is not None else "X: -, Y: -")
    else:
        coord_var.set("X: -, Y: -")

canvas.mpl_connect('motion_notify_event', on_mouse_move)

# Dialog Windows
dialog_windows = {}


def toggle_dialog(graph):
    if dialog_vars[graph].get():
        if graph not in dialog_windows:
            dialog = Toplevel(root)
            dialog.title(f"{graph.capitalize()} Graph")
            screen_width = root.winfo_screenwidth()
            screen_height = root.winfo_screenheight()
            width = screen_width # half width
            height = screen_height # half height
            dialog.geometry(f"{width//4}x{height//2}")
            dialog.protocol("WM_DELETE_WINDOW", lambda g=graph: close_dialog(g))

            fig_dialog, ax_dialog = plt.subplots(figsize=(8, 6))
            ax_dialog.set_xlabel('Time')
            ax_dialog.set_ylabel(get_label(graph))
            ax_dialog.grid(True)
            line_dialog, = ax_dialog.plot([], [], get_color(graph), label=graph.capitalize())
            ax_dialog.legend(loc='upper left')
            ax_dialog.set_ylim(get_ylim(graph))
            fig_dialog.autofmt_xdate()
            fig_dialog.tight_layout()

            canvas_dialog = FigureCanvasTkAgg(fig_dialog, master=dialog)
            canvas.draw()
            canvas_dialog.get_tk_widget().pack(fill="both", expand=True)
            toolbar = NavigationToolbar2Tk(canvas, dialog)
            toolbar.update()
            toolbar.pack(anchor='n', fill=tk.X)

            dialog_windows[graph] = {'window': dialog, 'fig': fig_dialog, 'ax': ax_dialog, 'line': line_dialog, 'canvas': canvas_dialog}
            ani_dialog = animation.FuncAnimation(fig_dialog, lambda i, g=graph: update_dialog_graph(i, g), interval=PLOT_UPDATE_INTERVAL_MS, cache_frame_data=False)
            dialog_windows[graph]['ani'] = ani_dialog
    else:
        close_dialog(graph)

def close_dialog(graph):
    if graph in dialog_windows:
        dialog_windows[graph]['ani'].event_source.stop()
        dialog_windows[graph]['window'].destroy()
        del dialog_windows[graph]
        dialog_vars[graph].set(False)

def update_dialog_graph(i, graph):
    if graph in dialog_windows and not panning:
        data_map = {'temperature': plot_temperatures, 'humidity': plot_humidities, 'temp2': plot_temp2,
                    'humidity2': plot_humidity2, 'heat_index': plot_heat_index, 'heat_index2': plot_heat_index2}
        if plot_timestamps:
            dialog_windows[graph]['line'].set_data(list(plot_timestamps), list(data_map[graph]))
            if len(plot_timestamps) > 1:
                dialog_windows[graph]['ax'].set_xlim(min(plot_timestamps), max(plot_timestamps))
            dialog_windows[graph]['ax'].relim()
            dialog_windows[graph]['ax'].autoscale_view(scalex=False)
            dialog_windows[graph]['canvas'].draw()

def set_graph(graph):
    global active_graph
    active_graph = graph
    update_graph_label()

def get_label(graph):
    return {'temperature': 'Temperature (°C)', 'humidity': 'Humidity (%)', 'temp2': 'Temp2 (°C)',
            'humidity2': 'Humidity2 (%)', 'heat_index': 'Heat Index (°C)', 'heat_index2': 'Heat Index2 (°C)'}[graph]

def get_ylim(graph):
    return {'temperature': (20, 60), 'humidity': (40, 100), 'temp2': (-10, 60), 'humidity2': (0, 100),
            'heat_index': (20, 60), 'heat_index2': (0, 60)}[graph]

def get_color(graph):
    return {'temperature': 'r', 'humidity': 'b', 'temp2': 'g', 'humidity2': 'm', 'heat_index': 'orange', 'heat_index2': 'purple'}[graph]

def update_graph_label():
    ax.set_ylabel(get_label(active_graph))
    ax.set_ylim(get_ylim(active_graph))
    line.set_color(get_color(active_graph))
    line.set_label(get_label(active_graph))
    ax.legend()
    canvas.draw()

def browse_directory(save_dir_var):
    global SAVE_DIR
    new_dir = filedialog.askdirectory(initialdir=SAVE_DIR, title="Select Save Directory")
    if new_dir:
        SAVE_DIR = new_dir
        save_dir_var.set(SAVE_DIR)

def start_logging():
    global ser, new_records_buffer, plot_timestamps, plot_temperatures, plot_humidities, plot_temp2, plot_humidity2
    global plot_heat_index, plot_heat_index2, last_save_time, stop_event, data_queue, ani, COM_PORT, BAUD_RATE
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

    # Check if COM port is connected
    try:
        ser = serial.Serial(COM_PORT, BAUD_RATE, timeout=1)
        ser.close()  # Close immediately after checking to avoid holding the port
    except serial.SerialException as se:
        messagebox.showerror("Serial Error", f"COM port {COM_PORT} is not available or in use: {se}")
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

    # Start animation
    ani = animation.FuncAnimation(fig, animate, interval=PLOT_UPDATE_INTERVAL_MS, cache_frame_data=False)
    canvas.draw()

    # Update UI
    start_button.config(state="disabled")
    calibrate_button.config(state="")
    stop_button.config(state="normal")
    status_var.set("Logging started")
    messagebox.showinfo("Info", "Logging started.")

def stop_logging():
    global ser, ani, stop_event
    stop_event.set()
    
    if ani:
        ani.event_source.stop()
    
    # Fix: Ensure ser is valid before calling close
    if ser is not None:
        try:
            if ser.is_open:
                ser.close()
                status_var.set("Serial port closed")
        except Exception as e:
            status_var.set(f"Error closing serial port: {e}")
    
    if new_records_buffer:
        status_var.set(f"Saving {len(new_records_buffer)} remaining records...")
        save_to_excel(new_records_buffer)
        new_records_buffer.clear()
    
    start_button.config(state="normal")
    stop_button.config(state="disabled")
    status_var.set("Logging stopped")
    messagebox.showinfo("Info", "Logging stopped.")


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
            if line:
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
                            'temp': temp, 'rh': rh, 'temp2': temp2,
                            'humidity2': humidity2, 'heat_index': heat_index,
                            'heat_index2': heat_index2
                        })
                        status_var.set(f"Logged: {device_time_str}, T1={temp}°C, H1={heat_index}°C")
                    except (ValueError, IndexError) as e:
                        status_var.set(f"Data parsing error: {e}")
        except serial.SerialException as se:
            status_var.set(f"Serial error: {se}. Attempting reconnect...")
            time.sleep(1)
        except Exception as e:
            status_var.set(f"Unexpected error: {e}")

    if ser and ser.is_open:
        ser.close()
        status_var.set("Serial port closed")

def animate(i):
    global last_save_time, new_records_buffer, latest_values
    while not data_queue.empty():
        data = data_queue.get()
        try:
            dt_object = datetime.strptime(data['timestamp_str'], '%d/%m/%Y %H:%M:%S')
        except ValueError as ve:
            status_var.set(f"Timestamp parse error: {ve}")
            return

        plot_timestamps.append(dt_object)
        plot_temperatures.append(data['temp'])
        plot_humidities.append(data['rh'])
        plot_temp2.append(data['temp2'])
        plot_humidity2.append(data['humidity2'])
        plot_heat_index.append(data['heat_index'])
        plot_heat_index2.append(data['heat_index2'])

        latest_values.update(data)
        update_real_time_values()

        new_records_buffer.append([data['timestamp_str'], data['temp'], data['rh'],
                                 data['temp2'], data['humidity2'], data['heat_index'],
                                 data['heat_index2']])

    if plot_timestamps and not panning:
        data_map = {
            'temperature': plot_temperatures, 'humidity': plot_humidities,
            'temp2': plot_temp2, 'humidity2': plot_humidity2,
            'heat_index': plot_heat_index, 'heat_index2': plot_heat_index2
        }
        line.set_data(list(plot_timestamps), list(data_map[active_graph]))
        if len(plot_timestamps) > 1:
            ax.set_xlim(min(plot_timestamps), max(plot_timestamps))
        ax.relim()
        ax.autoscale_view(scalex=False)
        canvas.draw()

    current_time = time.time()
    if (len(new_records_buffer) >= SAVE_INTERVAL_RECORDS) or \
       (new_records_buffer and (current_time - last_save_time) >= SAVE_INTERVAL_SECONDS):
        save_to_excel(new_records_buffer)
        new_records_buffer.clear()
        last_save_time = current_time

def update_real_time_values():
    temp1_value_label.config(text=f"{latest_values['temp']:.2f} °C")
    temp1_value_label.config(background="yellow")
    root.after(100, lambda: temp1_value_label.config(background="SystemButtonFace"))

    heat_index1_value_label.config(text=f"{latest_values['heat_index']:.2f} °C")
    heat_index1_value_label.config(background="yellow", foreground="green" if abs(latest_values['heat_index'] - latest_values['temp']) <= 5 else "red")
    root.after(100, lambda: heat_index1_value_label.config(background="SystemButtonFace"))

    temp2_value_label.config(text=f"{latest_values['temp2']:.2f} °C")
    temp2_value_label.config(background="yellow")
    root.after(100, lambda: temp2_value_label.config(background="SystemButtonFace"))

    heat_index2_value_label.config(text=f"{latest_values['heat_index2']:.2f} °C")
    heat_index2_value_label.config(background="yellow", foreground="green" if abs(latest_values['heat_index2'] - latest_values['temp2']) <= 5 else "red")
    root.after(100, lambda: heat_index2_value_label.config(background="SystemButtonFace"))

def save_to_excel(records):
    date_str = datetime.now().strftime("%Y%m%d")
    excel_file = os.path.join(SAVE_DIR, f"fluke_1620A_{date_str}.xlsx")
    temp_file = os.path.join(SAVE_DIR, f"fluke_1620A_{date_str}_tmp.xlsx")
    if not records:
        return
    temp_df = pd.DataFrame(records, columns=columns)
    try:
        if os.path.exists(excel_file):
            existing_df = pd.read_excel(excel_file, engine='openpyxl')
            updated_df = pd.concat([existing_df, temp_df], ignore_index=True)
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
            os.remove(temp_file)

def calibrate_time():
    global ser, stop_event

    if ser is None or not ser.is_open:
        messagebox.showerror("Error", "Serial port not open.")
        return

    try:
        now = datetime.now()

        date_cmd = f"SYST:DATE {now.year},{now.month},{now.day}\n"
        time_cmd = f"SYST:TIME {now.hour},{now.minute},{now.second}\n"

        ser.write(date_cmd.encode())
        time.sleep(0.5)
        ser.write(time_cmd.encode())
        time.sleep(0.5)

        msg = (
            f"Calibration complete:\n"
            f"Sent Date: {date_cmd[9:].strip().replace(',','/')}\n"
            f"Sent Time: {time_cmd[9:].strip().replace(',','/')}\n"
        )

        messagebox.showinfo("Time Calibration", msg)
        status_var.set("Time calibration completed")

    except Exception as e:
        messagebox.showerror("Error", f"Time calibration failed: {e}")
        status_var.set("Time calibration error")

# --- Start GUI ---
root.mainloop()
