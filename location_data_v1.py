import pandas as pd
import simplekml
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar, Combobox
from tkcalendar import DateEntry
import os
import threading

def convert_timestamp(ts):
    try:
        # iPhone epoch starts from 2001-01-01
        iphone_epoch_start = datetime(2001, 1, 1)
        utc_time = iphone_epoch_start + timedelta(seconds=float(ts))
        brisbane_time = utc_time + timedelta(hours=10)
        return brisbane_time.strftime('%d/%m/%Y'), brisbane_time.strftime('%H:%M:%S'), 'AEST (UTC+10)'
    except ValueError:
        return ts, ts, 'Unknown'  # Return as-is if conversion fails

def log_message(message):
    log_window.insert(tk.END, message + "\n")
    log_window.see(tk.END)

def process_file(excel_path, output_folder, start_datetime, end_datetime, horizontal_accuracy_filter, progress_bar, show_date, show_time, show_speed, show_bearing):
    try:
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
        elif horizontal_accuracy_filter == "< 50m":
            df = df[df["ZHORIZONTALACCURACY"] < 50]
        elif horizontal_accuracy_filter == "< 100m":
            df = df[df["ZHORIZONTALACCURACY"] < 100]
        elif horizontal_accuracy_filter == "< 500m":
            df = df[df["ZHORIZONTALACCURACY"] < 500]

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
            speed_mps = row.get("ZSPEED", 0)  # Speed in meters per second
            speed_kmh = round(speed_mps * 3.6, 1)  # Convert speed from m/s to km/h and round to 1 decimal place
            course = row.get("ZCOURSE", "No data recorded")
            if course == -1:
                course = "No data recorded"
            if speed_mps == -1:
                speed_text = "No data recorded"
            else:
                speed_text = f"{speed_kmh} km/h ({speed_mps} m/s)"

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
        output_filename = f"Exported - {os.path.splitext(input_filename)[0]}.kml"
        output_kml = os.path.join(output_folder, output_filename)
        log_message(f"Saving KML file to: {output_kml}")
        kml.save(output_kml)
        log_message(f"KML file created: {output_kml}")
        log_message(f"Total data points created: {point_count}")
        messagebox.showinfo("Success", f"KML file created: {output_kml}\nTotal data points created: {point_count}")

    except Exception as e:
        log_message(f"An error occurred: {e}")
        messagebox.showerror("Error", f"An error occurred: {e}")

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        excel_path_entry.delete(0, tk.END)
        excel_path_entry.insert(0, file_path)

def browse_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        output_folder_entry.delete(0, tk.END)
        output_folder_entry.insert(0, folder_path)

def update_date_label(entry, label):
    date = entry.get_date()
    formatted_date = date.strftime('%A, %d %B %Y')
    label.config(text=formatted_date)

def run():
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

    # Reset background colors
    excel_path_entry.config(bg="white")
    output_folder_entry.config(bg="white")
    start_time_entry.config(bg="white")
    end_time_entry.config(bg="white")

    # Check for empty inputs and highlight in red if empty
    if not excel_path:
        excel_path_entry.config(bg="red")
    if not output_folder:
        output_folder_entry.config(bg="red")
    if not start_time:
        start_time_entry.config(bg="red")
    if not end_time:
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
    threading.Thread(target=process_file, args=(excel_path, output_folder, start_datetime, end_datetime, horizontal_accuracy_filter, progress_bar, show_date, show_time, show_speed, show_bearing)).start()

def validate_time_format(time_str):
    try:
        datetime.strptime(time_str, "%H:%M")
        return True
    except ValueError:
        return False

# Create the main window
root = tk.Tk()
root.title("IPhone Location Data Map Exporter")

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
tk.Checkbutton(root, text="Speed", variable=speed_var).grid(row=4, column=2, padx=10, pady=5, sticky="w")
tk.Checkbutton(root, text="Bearing", variable=bearing_var).grid(row=4, column=3, padx=10, pady=5, sticky="w")

tk.Label(root, text="Start Date:").grid(row=5, column=2, padx=10, pady=10, sticky="e")
start_date_entry = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
start_date_entry.grid(row=5, column=3, padx=10, pady=10, sticky="w")
start_date_label = tk.Label(root, text="")
start_date_label.grid(row=5, column=4, padx=10, pady=10, sticky="w")
start_date_entry.bind("<<DateEntrySelected>>", lambda event: update_date_label(start_date_entry, start_date_label))

tk.Label(root, text="Start Time (HH:MM) 24hr:").grid(row=5, column=5, padx=10, pady=10, sticky="e")
start_time_entry = tk.Entry(root, width=10)
start_time_entry.grid(row=5, column=6, padx=10, pady=10, sticky="w")

tk.Label(root, text="End Date:").grid(row=6, column=2, padx=10, pady=10, sticky="e")
end_date_entry = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
end_date_entry.grid(row=6, column=3, padx=10, pady=10, sticky="w")
end_date_label = tk.Label(root, text="")
end_date_label.grid(row=6, column=4, padx=10, pady=10, sticky="w")
end_date_entry.bind("<<DateEntrySelected>>", lambda event: update_date_label(end_date_entry, end_date_label))

tk.Label(root, text="End Time (HH:MM) 24hr:").grid(row=6, column=5, padx=10, pady=10, sticky="e")
end_time_entry = tk.Entry(root, width=10)
end_time_entry.grid(row=6, column=6, padx=10, pady=10, sticky="w")

tk.Label(root, text="Horizontal Accuracy:").grid(row=7, column=0, padx=10, pady=10, sticky="e")
horizontal_accuracy_combobox = Combobox(root, values=["nil", "< 10m", "< 50m", "< 100m", "< 500m"], state="readonly")
horizontal_accuracy_combobox.grid(row=7, column=1, padx=10, pady=10, sticky="w")
horizontal_accuracy_combobox.current(0)  # Set default value to "nil"

tk.Button(root, text="Run", command=run, width=20, height=2).grid(row=8, column=0, columnspan=5, padx=10, pady=20)

progress_bar = Progressbar(root, orient="horizontal", length=400, mode="determinate")
progress_bar.grid(row=9, column=0, columnspan=5, padx=10, pady=10)

# Create the log window
log_window = tk.Text(root, height=10, width=80)
log_window.grid(row=10, column=0, columnspan=5, padx=10, pady=10)

# Run the application
root.mainloop()