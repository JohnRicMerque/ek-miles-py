import requests
import pandas as pd
import time

def calculate_miles(from_airport, to_airport, dep_date, ret_date, dep_class, ret_class, adults, ofws, children, infants, teenagers, by, travel_type, tier):
    url = f"https://www.emirates.com/service/ekl/loyalty/calculate-miles?airline=EK&origin={from_airport}&destination={to_airport}&cabin={dep_class}&journeyType=OW&tier={tier}"

    # Update the cookie dynamically with provided parameters
    cookie_value = f'ps={{"From":"{from_airport}","To":"{to_airport}","DepDate":"{dep_date}","RetDate":"{ret_date}","DepClass":"{dep_class}","RetClass":"{ret_class}","Adults":"{adults}","OFWs":"{ofws}","Child":"{children}","Infants":"{infants}","Teenager":"{teenagers}","By":"{by}","Type":"{travel_type}"}};'

    headers = {
        'accept': 'application/json, text/plain, */*',
        'accept-language': 'en-US,en;q=0.9',
        'cookie': cookie_value,
        'priority': 'u=1, i',
        'referer': 'https://www.emirates.com/ph/english/skywards/miles-calculator/',
        'sec-ch-ua': '"Chromium";v="128", "Not;A=Brand";v="24", "Google Chrome";v="128"',
        'sec-ch-ua-mobile': '?1',
        'sec-ch-ua-platform': '"Android"',
        'sec-fetch-dest': 'empty',
        'sec-fetch-mode': 'cors',
        'sec-fetch-site': 'same-origin',
        'user-agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Mobile Safari/537.36'
    }

    response = requests.get(url, headers=headers)
    return response.json()

# Load the Excel file (replace with your actual Excel file path)
file_path = 'input/input.xlsx'

# Read the Excel file into a DataFrame
df = pd.read_excel(file_path)

rows = []
# Loop through each row in the DataFrame
for index, row in df.iterrows():
    time.sleep(2)
    # Accessing each column's value in the current row
    one_way_or_roundtrip = row['oneWayOrRoundtrip']
    flying_with = row['flyingWith']
    leaving_from = row['leavingFrom']
    going_to = row['goingTo']
    cabin_class = row['cabinClass']
    emirates_skywards_tier = row['emiratesSkywardsTier']

    print(f"{leaving_from} - {going_to} - {cabin_class} - {emirates_skywards_tier} - Scraping...")
    #Y for economy, J for business, F for first class, W premium economy
    match cabin_class:
        case 'Economy':
            dep_class = 'Y'
        case 'Business':
            dep_class = 'J'
        case 'First':
            dep_class = 'F'
        case 'Premium Economy':
            dep_class = 'W'
        case _:  # default case
            dep_class = 'Unknown'

    # Constant Parameters Except Date
    ret_class = "7"
    dep_date = "100924"  # format: DDMMYY
    ret_date = "151024"  # format: DDMMYY# 
    adults = "2"
    ofws = "0"
    children = "0"
    infants = "0"
    teenagers = "0"
    by = "0"
    travel_type = "0"

    miles_data = calculate_miles(leaving_from, going_to, dep_date, ret_date, dep_class, ret_class, adults, ofws, children, infants, teenagers, by, travel_type, emirates_skywards_tier.lower())

    # print(miles_data)

    # Extracting the relevant data
    origin = miles_data['getMilesFromCouchbase']['origin']
    destination = miles_data['getMilesFromCouchbase']['destination']
    cabin = miles_data['getMilesFromCouchbase']['cabin']
    journey_type = miles_data['getMilesFromCouchbase']['journeyType']

    if emirates_skywards_tier.lower() == 'blue':
        tier = 'skywards'
    else:
        tier = emirates_skywards_tier.lower()
    
    earn_miles = miles_data['getMilesFromCouchbase']['miles']['earn'][tier]
    # Structuring the data for Excel
    fares = [ 'flexPlus','flex', 'saver', 'special']

    for fare in fares:
        if dep_class == 'F' and fare in ['saver', 'special']:
            continue  # skip this iteration

        skywards_miles = (earn_miles.get(fare, {}).get('skywardsMiles') or 'N/A')
        tier_miles = (earn_miles.get(fare, {}).get('tierMiles') or 'N/A')

        # Format data
        match cabin:
            case 'Y':
                cabin_formatted = 'Economy'
            case 'J':
                cabin_formatted = 'Business'
            case 'F':
                cabin_formatted = 'First'
            case 'W':
                cabin_formatted = 'Premium Economy'
            case _:  # default case
                cabin_formatted = 'Unknown'

        match journey_type:
            case 'OW':
                journey_type_formatted = 'One Way'
            case 'RT':
                journey_type_formatted = 'Round Trip'
            case _:  # default case
                journey_type_formatted = 'Unknown'
        
        if cabin_formatted == "Premium Economy":
            formatted_fare = f"{fare.capitalize()}"
        else:
            formatted_fare = f"{cabin_formatted} {fare.capitalize()}"

        if str(skywards_miles).isdigit():
            skywards_miles = int(skywards_miles)
        else:
            skywards_miles = skywards_miles

        if str(tier_miles).isdigit():
            tier_miles = int(tier_miles)
        else:
            tier_miles = tier_miles

        row = {
            'Direction': journey_type_formatted,
            'Airline': 'Emirates',
            'Leaving from': origin,
            'Going to': destination,
            'Cabin Class': cabin_formatted,
            'Skywards Tier': emirates_skywards_tier.capitalize(),
            'Fare':formatted_fare,
            'Skywards Miles': skywards_miles,
            'Tier Miles': tier_miles
        }
        rows.append(row)

    print(f"{leaving_from} - {going_to} - {cabin_class} - {emirates_skywards_tier} - Scraped Successfully!")
# Creating a DataFrame
df = pd.DataFrame(rows)

# Saving the DataFrame to an Excel file
output_file = 'data/skywards_miles_data.xlsx'
df.to_excel(output_file, index=False)

print(f"Data saved to {output_file}")
