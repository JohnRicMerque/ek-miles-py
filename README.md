# ✈️ Frequent Flyer Miles Web Scraping Automation Tool

A Python-based automation tool that extracts mileage data from Emirates Airlines' public-facing **Miles Calculator** endpoints.  
This tool streamlines data collection by automating what would typically take months to complete manually, delivering results in just weeks.

---

## 🚀 Features
- **Automated Data Extraction**: Fetches frequent flyer mileage information directly from Emirates' endpoints.  
- **High Efficiency**: Handled **500,000+ rows** of mileage data in under 2 weeks, reducing a manual 2–3 month process.  
- **Excel Export with Formatting**: Outputs clean, structured data in `.xlsx` format with auto-adjusted rows, columns, and bold headers.  
- **Proxy Support**: Integrates rotating proxies for improved scraping reliability.  
- **User-Friendly GUI**: Simple Tkinter interface for file selection, scraping progress, and saving results.  

---

## 🛠️ Tech Stack
- **Python 3.10+**
- [Requests](https://docs.python-requests.org/) – API requests handling  
- [Pandas](https://pandas.pydata.org/) – Data processing  
- [OpenPyXL](https://openpyxl.readthedocs.io/) – Excel formatting and saving  
- [Tkinter](https://docs.python.org/3/library/tkinter.html) – GUI for file selection & progress tracking  
- [FreeProxy](https://pypi.org/project/free-proxy/) – Proxy rotation  

---

## ⚡ Installation
1. **Clone the repository**
   ```bash
   git clone https://github.com/your-username/frequent-flyer-miles-scraper.git
   cd frequent-flyer-miles-scraper
````

2. **Create and activate a virtual environment**

   ```bash
   python -m venv venv
   source venv/bin/activate   # Mac/Linux
   venv\Scripts\activate      # Windows
   ```

3. **Install dependencies**

   ```bash
   pip install -r requirements.txt
   ```

---

## ▶️ Usage

1. **Prepare input file**: Create an Excel file containing routes and parameters:

   * `oneWayOrRoundtrip`
   * `flyingWith`
   * `leavingFrom`
   * `goingTo`
   * `cabinClass`
   * `emiratesSkywardsTier`

2. **Run the tool**

   ```bash
   python miles_scraper.py
   ```

3. **Select input Excel file** when prompted via GUI.

4. **Click "Scrape"** to start data collection.

5. **Save results**: Output will be saved as an Excel file with the following fields:

   * Direction
   * Airline
   * Leaving from
   * Going to
   * Cabin Class
   * Skywards Tier
   * Fare
   * Skywards Miles
   * Tier Miles

---

## 📊 Sample Output (Excel)

| Direction | Airline  | Leaving from | Going to | Cabin Class | Skywards Tier | Fare          | Skywards Miles | Tier Miles |
| --------- | -------- | ------------ | -------- | ----------- | ------------- | ------------- | -------------- | ---------- |
| One Way   | Emirates | DXB          | LHR      | Economy     | Silver        | Economy Saver | 2250           | 1500       |
| One Way   | Emirates | DXB          | JFK      | Business    | Gold          | Business Flex | 25000          | 15000      |

---

## ⚠️ Disclaimer

This project is intended **for educational and research purposes only**.
Use responsibly and ensure compliance with the target website’s Terms of Service.

---

## 👨‍💻 Author

**John Ric Merque**
📍 Manila, Philippines
💼 [Portfolio](https://johnricmerque.vercel.app)

---

```

Would you like me to also **generate a `requirements.txt`** for this so your repo is plug-and-play, or keep dependencies listed only in the README?
```
