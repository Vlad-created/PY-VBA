import pandas as pd
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut
import time

# === CONFIGURATION ===
file_path = r"\\virgo.mobiltel.bg\Work\OpControl\2016\Vlad\Adhoc\Clients Stara Zagora\PID_Stara_Zagora July2025.xlsx"
sheet_name = "Sheet6"
address_column = "Customer address"  # Change if your column has a different name
output_file = "Geocoded_Results.xlsx"

# === LOAD DATA ===
df = pd.read_excel(file_path, sheet_name=sheet_name)

# === SETUP GEOCODER ===
geolocator = Nominatim(user_agent="geo-batch")

# === GEOCODING FUNCTION WITH RETRY ===
def geocode_address(address, retries=3):
    for attempt in range(retries):
        try:
            location = geolocator.geocode(address, timeout=10)
            if location:
                return pd.Series([location.latitude, location.longitude])
            else:
                return pd.Series([None, None])
        except GeocoderTimedOut:
            time.sleep(1)
    return pd.Series([None, None])

# === APPLY GEOCODING ===
print("Geocoding in progress...")
df[['Latitude', 'Longitude']] = df[address_column].apply(geocode_address)

# === SAVE RESULTS ===
df.to_excel(output_file, index=False)
print(f"Geocoding complete. Results saved to: {output_file}")
