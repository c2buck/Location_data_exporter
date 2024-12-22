# Ensure you have Pillow installed: pip install pillow
import pandas as pd
import simplekml
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel, Label, HORIZONTAL
from tkinter.ttk import Progressbar, Combobox
from tkcalendar import DateEntry
import os
import threading
from PIL import Image, ImageTk

def update_speed_unit_state():
    log_message("Updating speed unit state...")
    if speed_var.get():
        kmh_radiobutton.config(state=tk.NORMAL)
        ms_radiobutton.config(state=tk.NORMAL)
    else:
        kmh_radiobutton.config(state=tk.DISABLED)
        ms_radiobutton.config(state=tk.DISABLED)

def convert_timestamp(ts):
    try:
        log_message(f"Converting timestamp: {ts}")
        # iPhone epoch starts from 2001-01-01
        iphone_epoch_start = datetime(2001, 1, 1)
        utc_time = iphone_epoch_start + timedelta(seconds=float(ts))
        brisbane_time = utc_time + timedelta(hours=10)
        return brisbane_time.strftime('%d/%m/%Y'), brisbane_time.strftime('%H:%M:%S'), 'AEST (UTC+10)'
    except ValueError:
        log_message(f"Failed to convert timestamp: {ts}")
        return ts, ts, 'Unknown'  # Return as-is if conversion fails

def log_message(message):
    log_window.insert(tk.END, message + "\n")
    log_window.see(tk.END)

def process_file(excel_path, output_folder, start_datetime, end_datetime, horizontal_accuracy_filter, progress_bar, show_date, show_time, show_speed, show_bearing, speed_unit):
    try:
        log_message("Starting file processing...")
        log_message(f"Excel path: {excel_path}")
        log_message(f"Output folder: {output_folder}")
        log_message(f"Start datetime: {start_datetime}")
        log_message(f"End datetime: {end_datetime}")
        log_message(f"Horizontal accuracy filter: {horizontal_accuracy_filter}")
        log_message(f"Show date: {show_date}, Show time: {show_time}, Show speed: {show_speed}, Show bearing: {show_bearing}, Speed unit: {speed_unit}")

        # Define the column names
        column_names = [
            "Z_PK", "ZALTITUDE", "ZCOURSE", "ZHORIZONTALACCURACY", "ZLATITUDE", "ZLONGITUDE",
            "ZSPEED", "ZTIMESTAMP", "ZVERTICALACCURACY"
        ]

        # Read the Excel file into a pandas DataFrame with the correct column names
        log_message(f"Reading Excel file: {excel_path}")
        df = pd.read_excel(excel_path, usecols=column_names, dtype={
            "Z_PK": str, "ZALTITUDE": float, "ZCOURSE": float, "ZHORIZONTALACCURACY": float, "ZLATITUDE": float, "ZLONGITUDE": float,
            "ZSPEED": float, "ZTIMESTAMP": str, "ZVERTICALACCURACY": float
        })
        log_message("Excel file read successfully")

        # Check if latitude and longitude columns are present and not empty
        if "ZLATITUDE" not in df.columns or "ZLONGITUDE" not in df.columns:
            raise ValueError("Latitude or Longitude columns are missing in the file.")
        if df["ZLATITUDE"].isna().all() or df["ZLONGITUDE"].isna().all():
            raise ValueError("Latitude or Longitude columns are empty in the file.")

        # Filter the DataFrame based on the start and end datetime
        df["ZTIMESTAMP"] = df["ZTIMESTAMP"].astype(float)
        df["datetime"] = df["ZTIMESTAMP"].apply(lambda ts: datetime(2001, 1, 1) + timedelta(seconds=ts) + timedelta(hours=10))
        df = df[(df["datetime"] >= start_datetime) & (df["datetime"] <= end_datetime)]

        # Apply horizontal accuracy filter
        if horizontal_accuracy_filter == "< 10m":
            df = df[df["ZHORIZONTALACCURACY"] < 10]
            horizontal_accuracy_filter_str = "less than 10m"
        elif horizontal_accuracy_filter == "< 50m":
            df = df[df["ZHORIZONTALACCURACY"] < 50]
            horizontal_accuracy_filter_str = "less than 50m"
        elif horizontal_accuracy_filter == "< 100m":
            df = df[df["ZHORIZONTALACCURACY"] < 100]
            horizontal_accuracy_filter_str = "less than 100m"
        elif horizontal_accuracy_filter == "< 500m":
            df = df[df["ZHORIZONTALACCURACY"] < 500]
            horizontal_accuracy_filter_str = "less than 500m"
        else:
            horizontal_accuracy_filter_str = "nil"

        # Create a KML object
        log_message("Creating KML object...")
        kml = simplekml.Kml()

        # Define the style for the red dot icon
        red_dot_style = simplekml.Style()
        red_dot_style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png'
        red_dot_style.iconstyle.color = simplekml.Color.red  # Set the color to red
        red_dot_style.iconstyle.scale = 0.6  # Increase the icon size

        # Iterate through the DataFrame and create placemarks
        log_message("Iterating through DataFrame...")
        point_count = 0
        total_rows = len(df)
        for index, row in df.iterrows():
            lat = row["ZLATITUDE"]
            lon = row["ZLONGITUDE"]

            # Skip rows with missing latitude or longitude
            if pd.isna(lat) or pd.isna(lon):
                log_message(f"Skipping row with missing coordinates: lat={lat}, lon={lon}")
                continue

            alt = round(row.get("ZALTITUDE", 0), 1)  # Use altitude from the column and round to 1 decimal place
            vertical_accuracy = round(row.get("ZVERTICALACCURACY", 0), 1)  # Round to 1 decimal place
            horizontal_accuracy = round(row.get("ZHORIZONTALACCURACY", 0), 1)  # Round to 1 decimal place
            speed_mps = round(row.get("ZSPEED", 0), 1)  # Speed in meters per second, rounded to 1 decimal place
            speed_kmh = round(speed_mps * 3.6, 1)  # Convert speed from m/s to km/h and round to 1 decimal place
            course = round(row.get("ZCOURSE", 0), 1)  # Round course to 1 decimal place
            if course == -1:
                course = "No data recorded"
            if speed_mps == -1:
                speed_text = "No data recorded"
            else:
                if speed_unit == "km/h":
                    speed_text = f"{speed_kmh} km/h"
                else:
                    speed_text = f"{speed_mps} m/s"

            # Create a new point with the red dot style
            log_message(f"Creating point: coords: ({lon}, {lat}, {alt})")
            pnt = kml.newpoint(coords=[(lon, lat, alt)])
            pnt.style = red_dot_style

            # Convert timestamp to Brisbane time
            date_str, time_str, time_zone = convert_timestamp(row["ZTIMESTAMP"])

            # Set the description with the row data
            description = (
                "IPhone iOS location service Cache.sqlite-wal (Table: ZRTCLLOCATIONMO)\n"
                f"ID: {row['Z_PK']}\n"
                f"Time Zone: {time_zone}\n"
                f"Time: {time_str}\n"
                f"Date: {date_str}\n"
                f"Latitude: {lat}\n"
                f"Longitude: {lon}\n"
                f"Altitude: {alt} (m) radius\n"
                f"Vertical Accuracy: {vertical_accuracy} (m) radius\n"
                f"Horizontal Accuracy: {horizontal_accuracy} (m) radius\n"
                f"Course: {course}\n"
                f"Speed: {speed_text}"
            )
            pnt.description = description

            # Set the name with selected data points
            name_parts = []
            if show_date:
                name_parts.append(date_str)
            if show_time:
                name_parts.append(time_str)
            if show_speed:
                name_parts.append(speed_text)
            if show_bearing:
                name_parts.append(str(course))
            pnt.name = " | ".join(name_parts)

            point_count += 1

            # Update progress bar
            progress = int((index + 1) / total_rows * 100)
            progress_bar['value'] = progress
            root.update_idletasks()

        # Set the KML output file name
        input_filename = os.path.basename(excel_path)
        start_date_str = start_datetime.strftime('%Y%m%d%H%M')
        end_date_str = end_datetime.strftime('%Y%m%d%H%M')
        output_filename = f"Exported - {os.path.splitext(input_filename)[0]} - {horizontal_accuracy_filter_str} - {start_date_str}_to_{end_date_str}.kml"
        output_kml = os.path.join(output_folder, output_filename)
        log_message(f"Saving KML file to: {output_kml}")
        kml.save(output_kml)
        log_message(f"KML file created: {output_kml}")
        log_message(f"Total data points created: {point_count}")

        # Write filters and settings to a text file
        filters_filename = f"Filters - {os.path.splitext(input_filename)[0]}.txt"
        filters_path = os.path.join(output_folder, filters_filename)
        with open(filters_path, 'w') as f:
            f.write(f"Start Date: {start_datetime.strftime('%d/%m/%Y %H:%M')}\n")
            f.write(f"End Date: {end_datetime.strftime('%d/%m/%Y %H:%M')}\n")
            f.write(f"Horizontal Accuracy Filter: {horizontal_accuracy_filter}\n")
            f.write(f"Show Date: {show_date}\n")
            f.write(f"Show Time: {show_time}\n")
            f.write(f"Show Speed: {show_speed}\n")
            f.write(f"Show Bearing: {show_bearing}\n")
            f.write(f"Speed Unit: {speed_unit}\n")
        log_message(f"Filters and settings saved to: {filters_path}")

        # Display warning if more than 1000 data points
        if point_count > 1000:
            warning_message = (
                "WARNING - This file may crash Google Earth due to the large data volume."
                "Consider re-applying filters if there are issues."
            )
            log_message(f"WARNING: {warning_message}")
            messagebox.showwarning("Warning", warning_message)

        # Show success message with image
        show_success_message(output_kml, point_count, filters_path)

    except Exception as e:
        log_message(f"An error occurred: {e}")
        messagebox.showerror("Error", f"An error occurred: {e}")

def show_success_message(output_kml, point_count, filters_path):
    log_message("Showing success message...")
    success_window = Toplevel(root)
    success_window.title("Success")

    # Create and place the widgets
    Label(success_window, text=f"KML file created: {output_kml}").pack(pady=10)
    Label(success_window, text=f"Total data points created: {point_count}").pack(pady=5)
    Label(success_window, text=f"Filters and settings saved to: {filters_path}").pack(pady=5)

def browse_file():
    log_message("Browsing for file...")
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        excel_path_entry.delete(0, tk.END)
        excel_path_entry.insert(0, file_path)

def browse_folder():
    log_message("Browsing for folder...")
    folder_path = filedialog.askdirectory()
    if folder_path:
        output_folder_entry.delete(0, tk.END)
        output_folder_entry.insert(0, folder_path)

def update_date_label(entry, label):
    log_message("Updating date label...")
    date = entry.get_date()
    formatted_date = date.strftime('%A, %d %B %Y')
    label.config(text=formatted_date)

def run():
    log_message("Running process...")
    if not validate_speed_selection():
        return

    excel_path = excel_path_entry.get()
    output_folder = output_folder_entry.get()
    start_date = start_date_entry.get_date()
    start_time = start_time_entry.get()
    end_date = end_date_entry.get_date()
    end_time = end_time_entry.get()
    horizontal_accuracy_filter = horizontal_accuracy_combobox.get()

    # Get checkbox values
    show_date = date_var.get()
    show_time = time_var.get()
    show_speed = speed_var.get()
    show_bearing = bearing_var.get()
    speed_unit = speed_unit_var.get()

    # Reset background colors
    excel_path_entry.config(bg="white")
    output_folder_entry.config(bg="white")
    start_time_entry.config(bg="white")
    end_time_entry.config(bg="white")

    # Check for empty inputs and highlight in red if empty
    if not excel_path:
        log_message("Excel path is empty")
        excel_path_entry.config(bg="red")
    if not output_folder:
        log_message("Output folder is empty")
        output_folder_entry.config(bg="red")
    if not start_time:
        log_message("Start time is empty")
        start_time_entry.config(bg="red")
    if not end_time:
        log_message("End time is empty")
        end_time_entry.config(bg="red")

    if not excel_path or not output_folder or not start_time or not end_time:
        messagebox.showwarning("Input Error", "Please fill in all required fields.")
        return

    # Validate time format
    if not validate_time_format(start_time) or not validate_time_format(end_time):
        messagebox.showerror("Input Error", "Time must be in HH:MM format.")
        return

    start_datetime = datetime.combine(start_date, datetime.strptime(start_time, "%H:%M").time())
    end_datetime = datetime.combine(end_date, datetime.strptime(end_time, "%H:%M").time())
    threading.Thread(target=process_file, args=(excel_path, output_folder, start_datetime, end_datetime, horizontal_accuracy_filter, progress_bar, show_date, show_time, show_speed, show_bearing, speed_unit)).start()

def validate_time_format(time_str):
    log_message(f"Validating time format: {time_str}")
    try:
        datetime.strptime(time_str, "%H:%M")
        return True
    except ValueError:
        log_message(f"Invalid time format: {time_str}")
        return False

def show_warning():
    log_message("Showing warning message...")
    warning_message = (
        "WARNING\n"
        "Validate results independently before disclosing products in criminal proceedings\n"
        "Not to be disclosed outside of the QPS without prior written permission\n"
        "Application created by:\n"
        "XXX\n"
        "Please send email re bugs or feedback to:\n"
        "XXXX"
    )
    if not messagebox.askokcancel("Disclaimer", warning_message):
        root.destroy()

def validate_speed_selection():
    log_message("Validating speed selection...")
    if not speed_var.get() and (speed_unit_var.get() == "km/h" or speed_unit_var.get() == "m/s"):
        messagebox.showerror("Input Error", "Please select the Speed checkbox to enable speed unit selection.")
        return False
    return True

# Create the main window
root = tk.Tk()
root.title("IPhone Location Data Map Exporter v.0.1 Beta")

# Show warning message
root.after(0, show_warning)

# Create and place the widgets
tk.Label(root, text="Excel File:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
excel_path_entry = tk.Entry(root, width=50)
excel_path_entry.grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Browse...", command=browse_file).grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="Output Folder:").grid(row=1, column=0, padx=10, pady=10, sticky="e")
output_folder_entry = tk.Entry(root, width=50)
output_folder_entry.grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="Browse...", command=browse_folder).grid(row=1, column=2, padx=10, pady=10)

tk.Label(root, text="Time Zone AEST +10 UTC").grid(row=2, column=0, columnspan=5, padx=10, pady=10)

tk.Label(root, text="Filter Options", font=("Helvetica", 12, "bold", "underline")).grid(row=3, column=0, columnspan=5, padx=10, pady=10)

# Add checkboxes for additional data points
date_var = tk.BooleanVar()
time_var = tk.BooleanVar()
speed_var = tk.BooleanVar()
bearing_var = tk.BooleanVar()

tk.Checkbutton(root, text="Date", variable=date_var).grid(row=4, column=0, padx=10, pady=5, sticky="w")
tk.Checkbutton(root, text="Time", variable=time_var).grid(row=4, column=1, padx=10, pady=5, sticky="w")

tk.Checkbutton(root, text="Speed", variable=speed_var, command=update_speed_unit_state).grid(row=5, column=0, padx=10, pady=5, sticky="w")
tk.Checkbutton(root, text="Bearing", variable=bearing_var).grid(row=5, column=1, padx=10, pady=5, sticky="w")

# Add radio buttons for speed unit selection
speed_unit_var = tk.StringVar(value="km/h")
kmh_radiobutton = tk.Radiobutton(root, text="km/h", variable=speed_unit_var, value="km/h", state=tk.DISABLED)
kmh_radiobutton.grid(row=6, column=0, padx=10, pady=5, sticky="w")
ms_radiobutton = tk.Radiobutton(root, text="m/s", variable=speed_unit_var, value="m/s", state=tk.DISABLED)
ms_radiobutton.grid(row=6, column=1, padx=10, pady=5, sticky="w")

tk.Label(root, text="Start Date:").grid(row=7, column=0, padx=10, pady=10, sticky="e")
start_date_entry = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
start_date_entry.grid(row=7, column=1, padx=10, pady=10)
start_date_label = tk.Label(root, text="")
start_date_label.grid(row=7, column=2, padx=10, pady=10, sticky="w")
start_date_entry.bind("<<DateEntrySelected>>", lambda event: update_date_label(start_date_entry, start_date_label))

tk.Label(root, text="Start Time (HH:MM) 24hr:").grid(row=7, column=3, padx=10, pady=10, sticky="e")
start_time_entry = tk.Entry(root, width=10)
start_time_entry.grid(row=7, column=4, padx=10, pady=10, sticky="w")

tk.Label(root, text="End Date:").grid(row=8, column=0, padx=10, pady=10, sticky="e")
end_date_entry = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
end_date_entry.grid(row=8, column=1, padx=10, pady=10)
end_date_label = tk.Label(root, text="")
end_date_label.grid(row=8, column=2, padx=10, pady=10, sticky="w")
end_date_entry.bind("<<DateEntrySelected>>", lambda event: update_date_label(end_date_entry, end_date_label))

tk.Label(root, text="End Time (HH:MM) 24hr:").grid(row=8, column=3, padx=10, pady=10, sticky="e")
end_time_entry = tk.Entry(root, width=10)
end_time_entry.grid(row=8, column=4, padx=10, pady=10, sticky="w")

tk.Label(root, text="Horizontal Accuracy:").grid(row=9, column=0, padx=10, pady=10, sticky="e")
horizontal_accuracy_combobox = Combobox(root, values=["nil", "< 10m", "< 50m", "< 100m", "< 500m"], state="readonly")
horizontal_accuracy_combobox.grid(row=9, column=1, padx=10, pady=10, sticky="w")
horizontal_accuracy_combobox.current(0)  # Set default value to "nil"

tk.Button(root, text="Run", command=run, width=20, height=2).grid(row=10, column=0, columnspan=5, padx=10, pady=20)

progress_bar = Progressbar(root, orient=tk.HORIZONTAL, length=400, mode="determinate")
progress_bar.grid(row=11, column=0, columnspan=5, padx=10, pady=10)

# Create the log window
log_window = tk.Text(root, height=10, width=80)
log_window.grid(row=12, column=0, columnspan=5, padx=10, pady=10)

# Run the application
root.mainloop()