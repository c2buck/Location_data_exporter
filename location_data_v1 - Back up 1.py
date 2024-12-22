import pandas as pd
import simplekml
from datetime import datetime, timedelta

def main():
    # Set the Excel file path
    excel_path = r"C:\Users\micro\OneDrive\Documents\Location data Yohanna WAL.xlsx"

    try:
        # Define the column names
        column_names = [
            "Z_PK", "ZALTITUDE", "ZCOURSE", "ZHORIZONTALACCURACY", "ZLATITUDE", "ZLONGITUDE",
            "ZSPEED", "ZTIMESTAMP", "ZVERTICALACCURACY"
        ]

        # Read the Excel file into a pandas DataFrame with the correct column names
        print(f"Reading Excel file: {excel_path}")
        df = pd.read_excel(excel_path, usecols=column_names, dtype={
            "Z_PK": str, "ZALTITUDE": float, "ZCOURSE": float, "ZHORIZONTALACCURACY": float, "ZLATITUDE": float, "ZLONGITUDE": float,
            "ZSPEED": float, "ZTIMESTAMP": str, "ZVERTICALACCURACY": float
        })
        print("Excel file read successfully")

        # Inspect the columns
        print("DataFrame columns:", df.columns)
        print("First few rows of the DataFrame:")
        print(df.head(10))  # Inspect the first 10 rows

        # Check if latitude and longitude columns are present and not empty
        if "ZLATITUDE" not in df.columns or "ZLONGITUDE" not in df.columns:
            raise ValueError("Latitude or Longitude columns are missing in the file.")
        if df["ZLATITUDE"].isna().all() or df["ZLONGITUDE"].isna().all():
            raise ValueError("Latitude or Longitude columns are empty in the file.")

        # Convert timestamps from iPhone epoch time (seconds from UTC 2001) to Brisbane time (AEST, UTC+10)
        def convert_timestamp(ts):
            try:
                # iPhone epoch starts from 2001-01-01
                iphone_epoch_start = datetime(2001, 1, 1)
                utc_time = iphone_epoch_start + timedelta(seconds=float(ts))
                brisbane_time = utc_time + timedelta(hours=10)
                return brisbane_time.strftime('%d/%m/%Y'), brisbane_time.strftime('%H:%M:%S'), 'AEST (UTC+10)'
            except ValueError:
                return ts, ts, 'Unknown'  # Return as-is if conversion fails

        # Create a KML object
        print("Creating KML object...")
        kml = simplekml.Kml()

        # Define the style for the red dot icon
        red_dot_style = simplekml.Style()
        red_dot_style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png'
        red_dot_style.iconstyle.color = simplekml.Color.red  # Set the color to red
        red_dot_style.iconstyle.scale = 0.6  # Increase the icon size

        # Iterate through the DataFrame and create placemarks
        print("Iterating through DataFrame...")
        point_count = 0
        for _, row in df.iterrows():
            lat = row["ZLATITUDE"]
            lon = row["ZLONGITUDE"]

            # Skip rows with missing latitude or longitude
            if pd.isna(lat) or pd.isna(lon):
                print(f"Skipping row with missing coordinates: lat={lat}, lon={lon}")
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
            print(f"Creating point: coords: ({lon}, {lat}, {alt})")
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
            point_count += 1

        # Set the KML output file name
        output_kml = r"C:\Users\micro\OneDrive\Documents\locations.kml"
        print(f"Saving KML file to: {output_kml}")
        kml.save(output_kml)
        print(f"KML file created: {output_kml}")
        print(f"Total data points created: {point_count}")

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()