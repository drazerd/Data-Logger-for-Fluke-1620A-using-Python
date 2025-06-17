import serial
import pandas as pd
import time
from datetime import datetime
import matplotlib.pyplot as plt
import matplotlib.animation as animation
from collections import deque
import threading
import queue

# --- Configuration ---
COM_PORT = 'COM3'
BAUD_RATE = 9600
EXCEL_FILE = 'fluke_log_realtime.xlsx'

SAVE_INTERVAL_RECORDS = 60
SAVE_INTERVAL_SECONDS = 300

PLOT_MAX_POINTS = 300
PLOT_UPDATE_INTERVAL_MS = 1000

# --- Data Storage ---
columns = ['Device Timestamp', 'Temperature (°C)', 'Humidity (%)', 'Temp2 (°C)', 'Humidity2 (%)']
df = pd.DataFrame(columns=columns)
new_records_buffer = []

plot_timestamps = deque(maxlen=PLOT_MAX_POINTS)
plot_temperatures = deque(maxlen=PLOT_MAX_POINTS)
plot_humidities = deque(maxlen=PLOT_MAX_POINTS)
plot_temp2 = deque(maxlen=PLOT_MAX_POINTS)
plot_humidity2 = deque(maxlen=PLOT_MAX_POINTS)

last_save_time = time.time()

# --- Threading Setup ---
stop_event = threading.Event()
data_queue = queue.Queue()

# --- Matplotlib Setup ---
fig, (ax1, ax2, ax3, ax4) = plt.subplots(4, 1, sharex=True, figsize=(12, 10))
fig.suptitle('Real-Time Temperature, Humidity, Temp2, and Humidity2')

# Temperature
ax1.set_ylabel('Temperature (°C)')
ax1.grid(True)
line_temp, = ax1.plot([], [], 'r-', label='Temperature (°C)')
ax1.legend(loc='upper left')
ax1.set_ylim(20, 30)

# Humidity
ax2.set_ylabel('Humidity (%)')
ax2.grid(True)
line_rh, = ax2.plot([], [], 'b-', label='Humidity (%)')
ax2.legend(loc='upper left')
ax2.set_ylim(40, 60)

# Temp2
ax3.set_ylabel('Temp2 (°C)')
ax3.grid(True)
line_temp2, = ax3.plot([], [], 'g-', label='Temp2 (°C)')
ax3.legend(loc='upper left')
ax3.set_ylim(-10, 30)

# Humidity2
ax4.set_xlabel('Time')
ax4.set_ylabel('Humidity2 (%)')
ax4.grid(True)
line_humidity2, = ax4.plot([], [], 'm-', label='Humidity2 (%)')
ax4.legend(loc='upper left')
ax4.set_ylim(0, 100)

fig.autofmt_xdate()

# --- Serial Reader Thread ---
def serial_reader_thread():
    global ser, last_save_time, df

    ser = None
    try:
        ser = serial.Serial(COM_PORT, BAUD_RATE, timeout=1)
        print(f"Connected to {COM_PORT} at {BAUD_RATE} baud.")
    except serial.SerialException as se:
        print(f"Serial port error: {se}")
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

                    data_queue.put({
                        'timestamp_str': device_time_str,
                        'temp': temp,
                        'rh': rh,
                        'temp2': temp2,
                        'humidity2': humidity2
                    })

                except ValueError as ve:
                    print(f"Value error: {ve} in line: {line}")
                except IndexError as ie:
                    print(f"Index error: {ie} in line: {line}")

        except serial.SerialException as se:
            print(f"Serial port error: {se}. Attempting reconnect...")
            try:
                if ser and ser.is_open:
                    ser.close()
                ser = serial.Serial(COM_PORT, BAUD_RATE, timeout=1)
                print("Reconnected.")
            except Exception as re_e:
                print(f"Reconnect failed: {re_e}")
                time.sleep(1)

        except Exception as e:
            print(f"Unexpected error in reader: {e}")

    if ser and ser.is_open:
        ser.close()
        print("Serial port closed.")

# --- Animation Function ---
def animate(i):
    global last_save_time, df, new_records_buffer

    while not data_queue.empty():
        data = data_queue.get()
        dt_object = datetime.strptime(data['timestamp_str'], '%d/%m/%Y %H:%M:%S')

        plot_timestamps.append(dt_object)
        plot_temperatures.append(data['temp'])
        plot_humidities.append(data['rh'])
        plot_temp2.append(data['temp2'])
        plot_humidity2.append(data['humidity2'])

        new_records_buffer.append([
            data['timestamp_str'], data['temp'], data['rh'],
            data['temp2'], data['humidity2']
        ])

        print(f"Logged: {data['timestamp_str']}, T={data['temp']}°C, RH={data['rh']}%, T2={data['temp2']}°C, H2={data['humidity2']}%")

    if plot_timestamps:
        line_temp.set_data(plot_timestamps, plot_temperatures)
        line_rh.set_data(plot_timestamps, plot_humidities)
        line_temp2.set_data(plot_timestamps, plot_temp2)
        line_humidity2.set_data(plot_timestamps, plot_humidity2)

        xmin, xmax = min(plot_timestamps), max(plot_timestamps)
        ax1.set_xlim(xmin, xmax)
        ax3.set_xlim(xmin, xmax)
        ax4.set_xlim(xmin, xmax)

    current_time = time.time()
    if (len(new_records_buffer) >= SAVE_INTERVAL_RECORDS) or \
       (new_records_buffer and (current_time - last_save_time) >= SAVE_INTERVAL_SECONDS):
        print(f"Saving {len(new_records_buffer)} records to Excel...")
        temp_df = pd.DataFrame(new_records_buffer, columns=columns)
        df = pd.concat([df, temp_df], ignore_index=True)
        try:
            df.to_excel(EXCEL_FILE, index=False)
            print(f"Saved. Total records: {len(df)}")
            new_records_buffer.clear()
            last_save_time = current_time
        except Exception as e:
            print(f"Excel write error: {e}")

# --- Main ---
print("Starting real-time logger...")
print("Press Ctrl+C to stop.")

serial_thread = threading.Thread(target=serial_reader_thread, daemon=True)
serial_thread.start()

try:
    ani = animation.FuncAnimation(fig, animate, interval=PLOT_UPDATE_INTERVAL_MS, cache_frame_data=False)
    plt.show()

except KeyboardInterrupt:
    print("\n--- Interrupted by user ---")
except Exception as e:
    print(f"Main error: {e}")
finally:
    stop_event.set()
    serial_thread.join(timeout=5)

    if new_records_buffer:
        print(f"Saving final {len(new_records_buffer)} records...")
        temp_df = pd.DataFrame(new_records_buffer, columns=columns)
        df = pd.concat([df, temp_df], ignore_index=True)
        try:
            df.to_excel(EXCEL_FILE, index=False)
            print(f"Final save complete. Total records: {len(df)}")
        except Exception as e:
            print(f"Final save error: {e}")
    else:
        print("No new records to save.")

    print("Logger stopped.")
