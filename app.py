import streamlit as st
import pandas as pd
import simplekml
import polyline as pl
import requests
import os


# Function to merge coordinates
def merge_coordinates(stops_file, route_file, output_file):
    stops_df = pd.read_excel(stops_file)
    route_file_df = pd.ExcelFile(route_file)

    stops_df['Bus Stop'] = stops_df['Bus Stop'].str.title()

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for sheet_name in route_file_df.sheet_names:
            route_df = pd.read_excel(route_file_df, sheet_name=sheet_name)
            route_df['Bus Stop'] = route_df['Bus Stop'].str.title()
            route_updated_df = pd.merge(route_df, stops_df[['Bus Stop', 'center_lat', 'center_lon']], on='Bus Stop',
                                        how='left')
            route_updated_df = route_updated_df.drop_duplicates(subset=['Bus Stop'])
            route_updated_df.to_excel(writer, sheet_name=sheet_name, index=False)


# Function to decode polyline
def decode_polyline(encoded_polyline):
    return pl.decode(encoded_polyline)


# Function to get route coordinates from Google Maps API
def get_route_coordinates(origin, destination, api_key):
    url = f"https://maps.googleapis.com/maps/api/directions/json?origin={origin}&destination={destination}&key={api_key}"
    route_coordinates = []
    try:
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()
        if 'routes' in data and data['routes']:
            encoded_polyline = data['routes'][0]['overview_polyline']['points']
            decoded_points = pl.decode(encoded_polyline)
            route_coordinates = [(lat, lng) for lat, lng in decoded_points]
        return route_coordinates
    except requests.RequestException as e:
        st.error(f"Request failed: {e}")
        return []


# Function to create KML file
def create_kml(route_file, kml_file, api_key):
    df = pd.read_excel(route_file)

    if 'center_lat' not in df.columns or 'center_lon' not in df.columns:
        raise KeyError("Columns 'center_lat' and/or 'center_lon' not found in DataFrame")

    kml = simplekml.Kml()

    for i in range(len(df) - 1):
        origin = f"{df.loc[i, 'center_lat']},{df.loc[i, 'center_lon']}"
        destination = f"{df.loc[i + 1, 'center_lat']},{df.loc[i + 1, 'center_lon']}"
        route_coordinates = get_route_coordinates(origin, destination, api_key)
        if route_coordinates:
            linestring = kml.newlinestring(name=f"Route from {df.loc[i, 'Bus Stop']} to {df.loc[i + 1, 'Bus Stop']}")
            linestring.coords = [(lng, lat) for lat, lng in route_coordinates]
            linestring.style.linestyle.color = simplekml.Color.blue
            linestring.style.linestyle.width = 5

    kml.save(kml_file)


# Streamlit application
st.title("Bus Route KML Generator")

st.write("Upload the bus stops and bus routes Excel files to generate a KML file.")

stops_file = st.file_uploader("Upload Bus Stops File", type=["xlsx"])
route_file = st.file_uploader("Upload Bus Routes File", type=["xlsx"])
api_key = st.text_input("Use this if you dont have your API Key 'AIzaSyB9WZSBmm4pvLiHAfUFSnchnPtxRMrIVaU'")

if st.button("Generate KML"):
    if stops_file is not None and route_file is not None and api_key:
        stops_file_path = stops_file.name
        route_file_path = route_file.name
        merged_route_file_path = "merged_route.xlsx"

        with open(stops_file_path, "wb") as f:
            f.write(stops_file.getbuffer())

        with open(route_file_path, "wb") as f:
            f.write(route_file.getbuffer())

        try:
            merge_coordinates(stops_file_path, route_file_path, merged_route_file_path)
            kml_file_path = "bus_route_polyline.kml"
            create_kml(merged_route_file_path, kml_file_path, api_key)
            st.success("KML file generated successfully!")
            st.download_button("Download KML", open(kml_file_path, "rb"), file_name=kml_file_path)
        except KeyError as e:
            st.error(f"Error: {e}")
    else:
        st.error("Please upload both Excel files and enter the API key.")
