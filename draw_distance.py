import sqlite3
from geopy.geocoders import Nominatim
from geopy.distance import geodesic
import folium
import argparse
from dotenv import load_dotenv
import os

load_dotenv()
site_url = os.getenv('site_url')
file_url = os.getenv('file_url')
office_username = os.getenv('office_username')
password = os.getenv('password')

# db
db_password = os.getenv('db_password')
db_user = os.getenv('db_user')
db_server = os.getenv('db_server')
db_driver = os.getenv('db_driver')
db_db = os.getenv('db_db')

# Step 1: Fetch Data from Database
def fetch_shop_data(shop_codes):
    import pyodbc
    conn = pyodbc.connect(f'DRIVER={db_driver};'
                    f'SERVER={db_server};'
                    f'DATABASE={db_db};'
                    f'UID={db_user};'
                    f'PWD={db_password}')
    cursor = conn.cursor()
    placeholders = ', '.join('?' for _ in shop_codes)  # To match the number of shop codes
    query = f"SELECT code, city FROM contact_details_agent WHERE code IN ({placeholders})"
    cursor.execute(query, shop_codes)
    data = cursor.fetchall()
    conn.close()
    return data

# Step 2: Geocode the Cities
def geocode_cities(cities):
    geolocator = Nominatim(user_agent="shop_locator")
    city_coordinates = {}
    for city in cities:
        location = geolocator.geocode(city)
        if location:
            city_coordinates[city] = (location.latitude, location.longitude)
    return city_coordinates

# Step 3: Calculate Distances (Optional)
def calculate_distances(city_coordinates):
    distances = []
    cities = list(city_coordinates.keys())
    for i in range(len(cities)):
        for j in range(i + 1, len(cities)):
            dist = geodesic(city_coordinates[cities[i]], city_coordinates[cities[j]]).kilometers
            distances.append((cities[i], cities[j], dist))
    return distances

# Step 4: Plot on a Map
def plot_cities_on_map(city_coordinates, output_file):
    poland_map = folium.Map(location=[52.2297, 21.0122], zoom_start=6)
    for city, coords in city_coordinates.items():
        folium.Marker(location=coords, popup=city).add_to(poland_map)
    poland_map.save(output_file)

# Main Function
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Draw distances between shops on a map.")
    parser.add_argument("shop_codes", metavar="SHOP_CODE", type=str, nargs="+", help="List of shop codes")
    args = parser.parse_args()

    shop_codes = args.shop_codes  # Get shop codes from command line

    # Fetch data from the database
    shop_data = fetch_shop_data(shop_codes)

    # Fetch data from the database
    #shop_data = [(5, 'Kościan'), (10, 'Czacz'), (15, 'Leszno')]

    if not shop_data:
        print("No shops found for the provided codes.")
    else:
        # Extract city names from data
        cities = {city for _, city in shop_data}

        # Geocode cities
        city_coordinates = geocode_cities(cities)

        if not city_coordinates:
            print("Failed to geocode any cities.")
        else:
            # Optional: Calculate distances
            distances = calculate_distances(city_coordinates)
            for city1, city2, dist in distances:
                print(f"Distance between {city1} and {city2}: {dist:.2f} km")

            # Plot cities on the map
            output_file = "poland_shops_map.html"
            plot_cities_on_map(city_coordinates, output_file)
            print(f"Map saved as {output_file}. Open it in your browser to view.")
