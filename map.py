import pandas as pd
from geopy.geocoders import Nominatim
import folium
from folium.plugins import MarkerCluster
import time

# Load the Excel file
file_path = "/Users/ghazaleh/Documents/ECE88/ECE88-Reunion04.xlsx"
df = pd.read_excel(file_path)

# Clean and standardize the 'Current Location' column
df['Current Location'] = df['Current Location'].astype(str).str.strip().str.lower()

# Optional: Manual corrections for common duplicates
replace_dict = {
    'nyc': 'new york, ny',
    'new york city': 'new york, ny',
    'sf': 'san francisco, ca',
    's.f.': 'san francisco, ca',
    'la': 'los angeles, ca',
    'l.a.': 'los angeles, ca',
    'tehran, iran': 'tehran',
    # Add more if needed
}
df['Current Location'] = df['Current Location'].replace(replace_dict)

# Geocoding with caching to avoid redundant requests
geolocator = Nominatim(user_agent="ece88_mapper")
location_cache = {}

def get_coordinates(location):
    if location in location_cache:
        return location_cache[location]
    try:
        time.sleep(1)  # Avoid hitting API too fast
        loc = geolocator.geocode(location)
        if loc:
            coords = pd.Series([loc.latitude, loc.longitude])
            location_cache[location] = coords
            return coords
    except:
        pass
    return pd.Series([None, None])

# Apply geocoding
df[['Latitude', 'Longitude']] = df['Current Location'].apply(get_coordinates)

# Drop entries without coordinates
df = df.dropna(subset=['Latitude', 'Longitude'])

# Create base map centered on average coordinates
map_center = [df['Latitude'].mean(), df['Longitude'].mean()]
m = folium.Map(location=map_center, zoom_start=2)

# Use marker clustering
marker_cluster = MarkerCluster().add_to(m)

# Add each person to the map
for _, row in df.iterrows():
    name = f"{row['First Name']} {row['Last Name']}"
    position = row['Current Position'] if pd.notna(row['Current Position']) else "N/A"
    location = row['Current Location'].title()
    
    folium.Marker(
        location=[row['Latitude'], row['Longitude']],
        popup=f"<b>{name}</b><br>{position}<br>{location}"
    ).add_to(marker_cluster)

# Save map to HTML file
output_path = "/Users/ghazaleh/Documents/ECE88/ece88_reunion_map.html"
m.save(output_path)

print(f"Map saved to: {output_path}")
