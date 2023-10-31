import csv
import datetime
import openpyxl
import requests

def get_access_token():
    url = "https://api.baubuddy.de/index.php/login"
    payload = {
        "username": "365",
        "password": "1"
    }
    headers = {
        "Authorization": "Basic QVBJX0V4cGxvcmVyOjEyMzQ1NmlzQUxhbWVQYXNz",
        "Content-Type": "application/json"
    }
    response = requests.request("POST", url, json=payload, headers=headers)
    json_response = response.json()
    access_token = json_response["oauth"]["access_token"]
    return access_token

def get_vehicles_from_api(access_token):
    url = "https://api.baubuddy.de/dev/index.php/v1/vehicles/select/active"
    headers = {
        "Authorization": f"Bearer {access_token}"
    }
    response = requests.request("GET", url, headers=headers)
    json_response = response.json()
    vehicles = json_response["data"]
    return vehicles

def merge_vehicles_with_csv(vehicles, csv_path):
    with open(csv_path, "r") as csv_file:
        csv_reader = csv.reader(csv_file)
        csv_data = list(csv_reader)

    # Filter out any resources that do not have a value set for hu field
    vehicles = [vehicle for vehicle in vehicles if vehicle["hu"] is not None]

    # Resolve colorCode for each labelId
    for vehicle in vehicles:
        labelIds = vehicle["labelIds"]
        for labelId in labelIds:
            url = f"https://api.baubuddy.de/dev/index.php/v1/labels/{labelId}"
            headers = {
                "Authorization": f"Bearer {access_token}"
            }
            response = requests.request("GET", url, headers=headers)
            json_response = response.json()
            colorCode = json_response["data"]["colorCode"]
            vehicle["labelIds"][labelId] = {"colorCode": colorCode}

    return csv_data + vehicles

def generate_excel_file(vehicles, output_path, keys, colored=True):
    wb = openpyxl.Workbook()
    ws = wb.active

    # Write header row
    header = ["rnr"] + keys
    ws.append(header)

    # Write vehicle data
    for vehicle in vehicles:
        row = [vehicle["rnr"]]
        for key in keys:
            if key in vehicle:
                row.append(vehicle[key])
            else:
                row.append("")

        if colored:
            hu_date = datetime.datetime.strptime(vehicle["hu"], "%Y-%m-%d")
            today = datetime.datetime.today()
            delta = today - hu_date

            if delta.days <= 90:
                cell_color = "#007500"
            elif delta.days <= 365:
                cell_color = "#FFA500"
            else:
                cell_color = "#b30000"

            for cell in row:
                cell.fill = openpyxl.styles.PatternFill(fill_type="solid", start_color=cell_color)

        ws.append(row)

    wb.save(output_path)

def main():
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("-c", "--colored", action="store_true", default=True, help="Color rows based on HU expiration date")
    parser.add_
