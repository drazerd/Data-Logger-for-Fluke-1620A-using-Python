import serial
import pandas as pd
import time
from datetime import datetime
import matplotlib.pyplot as plt
import matplotlib.animation as animation
from collections import deque
import threading
import queue # For thread-safe data transfer

# --- Configuration ---
COM_PORT = 'COM20' 
BAUD_RATE = 9600
EXCEL_FILE = 'fluke_log_realtime.xlsx' 

# Save to Excel after every N new records or M seconds
SAVE_INTERVAL_RECORDS = 60 
SAVE_INTERVAL_SECONDS = 300 # 5 minutes

# Plotting configuration
PLOT_MAX_POINTS = 300 # Maximum number of data points to display on the plot at once
PLOT_UPDATE_INTERVAL_MS = 1000 # Plot refresh rate in milliseconds (e.g., 1000ms = 1 second)

# --- Data Storage ---
columns = ['Device Timestamp', 'Temperature (°C)', 'Humidity (%)']
df = pd.DataFrame(columns=columns) # Main DataFrame for all collected data
new_records_buffer = [] # Buffer for records before saving to Excel

# Deques for plotting - they automatically discard old data when maxlen is reached
plot_timestamps = deque(maxlen=PLOT_MAX_POINTS)
plot_temperatures = deque(maxlen=PLOT_MAX_POINTS)
plot_humidities = deque(maxlen=PLOT_MAX_POINTS)

last_save_time = time.time() # To track when the last save occurred

# --- Threading Event for graceful shutdown ---
stop_event = threading.Event() # Set this event to signal the worker thread to stop
data_queue = queue.Queue() # Thread-safe queue to pass data from serial reader to plotter

# --- Matplotlib Setup ---
fig, (ax1, ax2) = plt.subplots(2, 1, sharex=True, figsize=(10, 8))
fig.suptitle('Real-Time Temperature and Humidity')

ax1.set_ylabel('Temperature (°C)')
ax1.grid(True)
line_temp, = ax1.plot([], [], 'r-', label='Temperature (°C)')
ax1.legend(loc='upper left')

ax2.set_xlabel('Time')
ax2.set_ylabel('Humidity (%)')
ax2.grid(True)
line_rh, = ax2.plot([], [], 'b-', label='Humidity (%)')
ax2.legend(loc='upper left')

# Set up initial plot limits (can adjust dynamically later)
ax1.set_ylim(20, 30) # Example range for temperature
ax2.set_ylim(40, 60) # Example range for humidity

fig.autofmt_xdate() # Format x-axis labels for dates/times


# --- Worker Thread Function: Reads Serial Data ---
def serial_reader_thread():
    global ser, last_save_time, df

    ser = None # Initialize ser here for this thread's scope
    try:
        ser = serial.Serial(COM_PORT, BAUD_RATE, timeout=1)
        print(f"Serial reader thread connected to {COM_PORT} at {BAUD_RATE} baud.")
    except serial.SerialException as se:
        print(f"Serial port error in reader thread: {se}. Exiting thread.")
        stop_event.set() # Signal main thread to stop too
        return
    except Exception as e:
        print(f"An unexpected error occurred during serial setup in reader thread: {e}. Exiting thread.")
        stop_event.set()
        return

    while not stop_event.is_set(): # Keep running until stop_event is set
        try:
            line = ser.readline().decode(errors='ignore').strip()
            if not line:
                continue

            cleaned_line = line.replace('\xa0', ' ').strip()
            parts = [part.strip() for part in cleaned_line.split(',')]
            
            if len(parts) >= 4: 
                device_time_str = parts[0] 
                
                try:
                    temp = float(parts[1])          
                    rh = float(parts[3])  
                      

                    # Put processed data into the queue for the plotting thread
                    data_queue.put({'timestamp_str': device_time_str, 'temp': temp, 'rh': rh})
                    
                except ValueError as ve_conv:
                    print(f"Reader: Data conversion error for line '{line}': {ve_conv}")
                except IndexError as ie_idx:
                    print(f"Reader: Data parsing error (index out of bounds) for line '{line}': {ie_idx}")
                
            else:
                # print(f"Reader: Skipping line due to insufficient parts ({len(parts)} < 4): '{line}'") # Too verbose for continuous run
                pass # Silently skip malformed lines in real-time
        
        except serial.SerialException as se:
            print(f"Reader: Serial port error during read: {se}. Attempting to re-open...")
            if ser and ser.is_open:
                ser.close()
            ser = None # Mark for re-initialization
            time.sleep(1) # Wait a bit before retrying to connect
            try:
                ser = serial.Serial(COM_PORT, BAUD_RATE, timeout=1)
                print(f"Reader: Reconnected to {COM_PORT}.")
            except serial.SerialException as re_se:
                print(f"Reader: Failed to re-connect: {re_se}. Retrying next loop.")
                ser = None # Ensure it tries again
            except Exception as e_reconnect:
                print(f"Reader: Unexpected error during reconnect: {e_reconnect}. Retrying next loop.")
                ser = None

        except Exception as e:
            print(f"Reader: An unexpected error occurred: {e}")
            # Consider setting stop_event here if it's a critical error
            
    # Clean up when thread stops
    if ser and ser.is_open:
        ser.close()
        print("Serial reader thread: Serial port closed.")

# --- Animation Function: Updates Plot and Saves Data ---
def animate(i):
    global last_save_time, df, new_records_buffer # Declare global variables we'll modify

    # Process all available data from the queue
    while not data_queue.empty():
        data_point = data_queue.get()
        device_time_str = data_point['timestamp_str']
        temp = data_point['temp']
        rh = data_point['rh']

        # Convert timestamp string to datetime object for plotting
        dt_object = datetime.strptime(device_time_str, '%d/%m/%Y %H:%M:%S')
        
        # Add data to plotting deques
        plot_timestamps.append(dt_object)
        plot_temperatures.append(temp)
        plot_humidities.append(rh)

        # Add to buffer for Excel saving
        new_records_buffer.append([device_time_str, temp, rh])
        print(f"Plotter: Logged: {device_time_str}, T={temp}°C, RH={rh}% (Plot: {len(plot_timestamps)}, Buffer: {len(new_records_buffer)})")

    # Update plot data only if there's new data to show
    if plot_timestamps:
        line_temp.set_data(plot_timestamps, plot_temperatures)
        line_rh.set_data(plot_timestamps, plot_humidities)

        # Autoscale x-axis based on available data
        ax1.set_xlim(min(plot_timestamps), max(plot_timestamps))
    
    # --- Periodic Save Logic ---
    current_time = time.time()
    if (len(new_records_buffer) >= SAVE_INTERVAL_RECORDS) or \
       (new_records_buffer and (current_time - last_save_time) >= SAVE_INTERVAL_SECONDS):
        
        print(f"Plotter: Saving {len(new_records_buffer)} new records to Excel...")
        temp_df = pd.DataFrame(new_records_buffer, columns=columns)
        df = pd.concat([df, temp_df], ignore_index=True)
        
        try:
            df.to_excel(EXCEL_FILE, index=False)
            print(f"Plotter: Data saved successfully. Total records in Excel: {len(df)}")
            new_records_buffer.clear() # Clear the buffer after successful save
            last_save_time = current_time # Reset save time
        except Exception as excel_err:
            print(f"Plotter: Error saving to Excel: {excel_err}. Make sure the file isn't open.")
            # Do not clear buffer if save fails, try again next interval

# --- Main Execution Block ---
print("Starting real-time data logger with multi-threaded plot...")
print("Press Ctrl+C in the console to stop the program.")

serial_thread = threading.Thread(target=serial_reader_thread, daemon=True) # daemon=True means thread dies with main program
serial_thread.start()

try:
    # Set up the animation
    ani = animation.FuncAnimation(fig, animate, interval=PLOT_UPDATE_INTERVAL_MS, cache_frame_data=False)
    plt.show() # Display the plot and start the animation loop

except KeyboardInterrupt:
    print("\n--- Program interrupted by user (Ctrl+C) ---")
except Exception as e:
    print(f"An error occurred in the main plotting thread: {e}")
finally:
    # Signal the serial reading thread to stop
    stop_event.set() 
    serial_thread.join(timeout=5) # Wait for the thread to finish, with a timeout

    # Save any remaining data in the buffer before exiting
    if new_records_buffer:
        print(f"Final Save: Saving remaining {len(new_records_buffer)} records to Excel...")
        temp_df = pd.DataFrame(new_records_buffer, columns=columns)
        df = pd.concat([df, temp_df], ignore_index=True)
        try:
            df.to_excel(EXCEL_FILE, index=False)
            print(f"Final Save: Remaining data saved. Total records in Excel: {len(df)}")
        except Exception as excel_err:
            print(f"Final Save: Error saving final data to Excel: {excel_err}. File might be in use.")
    else:
        print("Final Save: No new records to save on exit.")

    print("Real-time data logger stopped.")