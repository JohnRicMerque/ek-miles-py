import requests
import pandas as pd

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

# Example usage with dynamic input
from_airport = "MNL"
to_airport = "ABJ"
tier='blue'
dep_class = "Y" #Y for economy, J for business, F for first class, W premium economy
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

miles_data = calculate_miles(from_airport, to_airport, dep_date, ret_date, dep_class, ret_class, adults, ofws, children, infants, teenagers, by, travel_type, tier)

# Extracting the relevant data
origin = miles_data['getMilesFromCouchbase']['origin']
destination = miles_data['getMilesFromCouchbase']['destination']
cabin = miles_data['getMilesFromCouchbase']['cabin']
journey_type = miles_data['getMilesFromCouchbase']['journeyType']
earn_miles = miles_data['getMilesFromCouchbase']['miles']['earn']['skywards']

# Structuring the data for Excel
fares = [ 'flexPlus','flex', 'saver', 'special']
rows = []

for fare in fares:
    skywards_miles = earn_miles[fare].get('skywardsMiles', 'N/A')
    tier_miles = earn_miles[fare].get('tierMiles', 'N/A')
    row = {
        'Direction': journey_type,
        'Airline': 'Emirates',
        'Leaving from': origin,
        'Going to': destination,
        'Cabin Class': cabin,
        'Skywards Tier': tier,
        'Fare': fare.capitalize(),
        'Skywards Miles': skywards_miles,
        'Tier Miles': tier_miles
    }
    rows.append(row)

# Creating a DataFrame
df = pd.DataFrame(rows)

# Saving the DataFrame to an Excel file
output_file = 'data/skywards_miles_data.xlsx'
df.to_excel(output_file, index=False)

print(f"Data saved to {output_file}")
