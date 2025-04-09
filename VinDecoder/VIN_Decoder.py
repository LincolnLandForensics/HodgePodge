import requests
from openpyxl import Workbook
import time

# Constants
VIN_API_ENDPOINT = "https://vpic.nhtsa.dot.gov/api/vehicles/DecodeVinValuesExtended/{vin}?format=json"
INPUT_FILE = "vins.txt"
OUTPUT_FILE = "vins.xlsx"

# Fields of interest from NHTSA response
FIELDS_TO_EXTRACT = [
    "VIN", "Make", "Model", "ModelYear", "BodyClass",
    "VehicleType", "EngineCylinders", "DisplacementL",
    "FuelTypePrimary", "TransmissionStyle", "PlantCountry"
]

def read_vins(file_path):
    """Reads VINs from a file, one per line."""
    with open(file_path, "r") as file:
        vins = [line.strip() for line in file if line.strip()]
    return vins

def decode_vin(vin):
    """Sends a request to NHTSA VIN decoder and extracts vehicle data."""
    url = VIN_API_ENDPOINT.format(vin=vin)
    response = requests.get(url)
    if response.status_code != 200:
        return {"VIN": vin, "Error": f"HTTP {response.status_code}"}
    
    data = response.json()
    if not data.get("Results"):
        return {"VIN": vin, "Error": "No result returned"}

    result = data["Results"][0]
    decoded = {field: result.get(field, "") for field in FIELDS_TO_EXTRACT}
    return decoded

def write_to_excel(data, output_path):
    """Writes list of dictionaries to an Excel file."""
    wb = Workbook()
    ws = wb.active
    ws.title = "VIN Data"

    # Write header
    ws.append(FIELDS_TO_EXTRACT)

    # Write data rows
    for entry in data:
        row = [entry.get(field, "") for field in FIELDS_TO_EXTRACT]
        ws.append(row)

    wb.save(output_path)

def main():
    vins = read_vins(INPUT_FILE)
    results = []
    for vin in vins:
        print(f"Decoding VIN: {vin}")
        try:
            decoded_data = decode_vin(vin)
            results.append(decoded_data)
        except Exception as e:
            print(f"Error decoding {vin}: {e}")
            results.append({"VIN": vin, "Error": str(e)})
        time.sleep(1)  # Respectful delay to avoid hammering the API

    write_to_excel(results, OUTPUT_FILE)
    print(f"VIN decoding complete. Results saved to {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
