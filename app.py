import streamlit as st
import pandas as pd
import simplekml
import polyline as pl
import requests
import os
import zipfile

# CSS styling
st.markdown(
    """
    <style>
    body {
        background-color: #f5f5f5;
        font-family: 'Arial', sans-serif;
    }
    .title {
        color: #2c3e50;
        text-align: center;
    }
    .file-upload, .api-key, .generate-btn {
        margin: 20px 0;
        text-align: center;
    }
    .file-upload input, .api-key input {
        width: 100%;
        padding: 10px;
        border: 1px solid #ccc;
        border-radius: 4px;
        box-sizing: border-box;
    }
    .generate-btn button {
        background-color: #3498db;
        color: white;
        border: none;
        padding: 10px 20px;
        text-align: center;
        text-decoration: none;
        display: inline-block;
        font-size: 16px;
        margin: 4px 2px;
        cursor: pointer;
        border-radius: 4px;
    }
    .generate-btn button:hover {
        background-color: #2980b9;
    }
    .success-msg, .error-msg {
        text-align: center;
        margin: 20px 0;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Function to merge coordinates
def merge_coordinates(stops_file, route_file, output_file):
    stops_df = pd.read_excel(stops_file)
    route_file_df = pd.ExcelFile(route_file)

    stops_df['Bus Stop'] = stops_df['Bus Stop'].str.title()

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for sheet_name in route_file_df.sheet_names:
            route_df = pd.read_excel(route_file_df, sheet_name=sheet_name)
            route_df['Bus Stop'] = route_df['Bus Stop'].str.title()
            route_updated_df = pd.merge(route_df, stops_df[['Bus Stop', 'center_lat', 'center_lon']], on='Bus Stop', how='left')
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

# Function to create KML files for each sheet
def create_kml(route_file, api_key, output_dir):
    route_file_df = pd.ExcelFile(route_file)
    kml_files = []

    for sheet_name in route_file_df.sheet_names:
        df = pd.read_excel(route_file_df, sheet_name=sheet_name)

        if 'center_lat' not in df.columns or 'center_lon' not in df.columns:
            raise KeyError("Columns 'center_lat' and/or 'center_lon' not found in DataFrame")

        kml = simplekml.Kml()
        all_route_coordinates = []

        for i in range(len(df) - 1):
            origin = f"{df.loc[i, 'center_lat']},{df.loc[i, 'center_lon']}"
            destination = f"{df.loc[i + 1, 'center_lat']},{df.loc[i + 1, 'center_lon']}"
            route_coordinates = get_route_coordinates(origin, destination, api_key)
            all_route_coordinates.extend(route_coordinates)

            if all_route_coordinates:
                linestring = kml.newlinestring(
                    name=f"Route from {df.loc[i, 'Bus Stop']} to {df.loc[i + 1, 'Bus Stop']}")
                linestring.coords = [(lng, lat) for lat, lng in all_route_coordinates]
                linestring.style.linestyle.color = simplekml.Color.blue
                linestring.style.linestyle.width = 5

        kml_file_name = os.path.join(output_dir, f"{sheet_name}.kml")
        kml.save(kml_file_name)
        kml_files.append(kml_file_name)

    return kml_files

# Streamlit application
st.title("Bus Route KML Generator", anchor="center", className="title")

st.write("Upload the bus stops and bus routes Excel files to generate KML files.")

stops_file = st.file_uploader("Upload Bus Stops File", type=["xlsx"], key="stops_file", className="file-upload")
route_file = st.file_uploader("Upload Bus Routes File", type=["xlsx"], key="route_file", className="file-upload")
api_key = st.text_input("Enter your Google Maps API Key. if you not have any, then use this 'AIzaSyB9WZSBmm4pvLiHAfUFSnchnPtxRMrIVaU'", key="api_key", className="api-key")

if st.button("Generate KMLs", className="generate-btn"):
    if stops_file is not None and route_file is not None and api_key:
        stops_file_path = stops_file.name
        route_file_path = route_file.name
        merged_route_file_path = "merged_route.xlsx"
        output_dir = "kml_files"

        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        with open(stops_file_path, "wb") as f:
            f.write(stops_file.getbuffer())

        with open(route_file_path, "wb") as f:
            f.write(route_file.getbuffer())

        try:
            merge_coordinates(stops_file_path, route_file_path, merged_route_file_path)
            kml_files = create_kml(merged_route_file_path, api_key, output_dir)

            # Create a ZIP file
            zip_file_path = "kml_files.zip"
            with zipfile.ZipFile(zip_file_path, 'w') as zipf:
                for kml_file in kml_files:
                    zipf.write(kml_file)

            st.success("KML files generated successfully!", className="success-msg")
            st.download_button("Download ZIP", open(zip_file_path, "rb"), file_name=zip_file_path)
        except KeyError as e:
            st.error(f"Error: {e}", className="error-msg")
    else:
        st.error("Please upload both Excel files and enter the API key.", className="error-msg")
