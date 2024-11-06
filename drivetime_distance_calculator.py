import pandas as pd
import googlemaps

# Replace with your Google Maps API key
API_KEY = 'AIzaSyDipSK_ZiMDYAkxPEK4gQa2YRQWA-V1jd8'
gmaps = googlemaps.Client(key=API_KEY)

# Load the Excel file
def load_excel(file_path):
    # Assuming the sheet has 'Site Address' and 'To Address' columns
    df = pd.read_excel(file_path)
    return df

# Function to calculate distance and drive time using Google Maps API
def get_distance_drive_time(start, end):
    try:
        result = gmaps.distance_matrix(start, end, mode="driving")
        distance = result['rows'][0]['elements'][0]['distance']['value']  # in meters
        duration = result['rows'][0]['elements'][0]['duration']['value']  # in seconds
        distance_km = distance / 1000  # convert to km
        duration_hours = duration / 3600  # convert to hours
        return distance_km, duration_hours
    except Exception as e:
        print(f"Error calculating for {start} to {end}: {e}")
        return None, None

# Main function to process the Excel file and update km, time, and round-trip values
def process_addresses(file_path, output_path):
    df = load_excel(file_path)

    # Add new columns for km, time, round trip km, and round trip time if not present
    if 'km' not in df.columns:
        df['km'] = ''
    if 'time' not in df.columns:
        df['time'] = ''
    if 'round_trip_km' not in df.columns:
        df['round_trip_km'] = ''
    if 'round_trip_time' not in df.columns:
        df['round_trip_time'] = ''

    for index, row in df.iterrows():
        start = row['Site Address']
        end = row['To Address']
        distance_km, duration_hours = get_distance_drive_time(start, end)
        if distance_km is not None and duration_hours is not None:
            df.at[index, 'km'] = distance_km
            df.at[index, 'time'] = duration_hours
            df.at[index, 'round_trip_km'] = distance_km * 2  # Multiply by 2 for round trip
            df.at[index, 'round_trip_time'] = duration_hours * 2  # Multiply by 2 for round trip

    # Save updated data back to Excel
    df.to_excel(output_path, index=False)
    print(f"File saved at {output_path}")

if __name__ == "__main__":
    input_file = 'https://municipalgroup-my.sharepoint.com/:x:/r/personal/stosoni_dexter_ca/Documents/Documents/Routing_File.csv?d=w21cfe4e5dfad43a5a49537d17888214c&csf=1&web=1&e=QerCTX'  # Full path to your input Excel file
    output_file = 'C:/Desktop/Routing_File_Done.csv'  # Full path for saving the output file
    process_addresses(input_file, output_file)
