# 1. Data Collection
# 2. Data Cleaning
# 3. Database Creation
# 4. Analytics
# 5. KPIs
# 6. Charts
# 7. Power BI Outputs


import pandas as pd
import numpy as np
import random
import os
from faker import Faker
from datetime import datetime, timedelta
import os
from openpyxl import Workbook

fake = Faker()

# Create folders if missing
os.makedirs("NordicSpace_Project/data", exist_ok=True)
num_rows = 5000

cities = {
    "Helsinki": {"base_price": 150},
    "Espoo": {"base_price": 135},
    "Vantaa": {"base_price": 125},
    "Tampere": {"base_price": 115},
    "Turku": {"base_price": 110}
}

unit_sizes = {
    "Small Locker": 0.6,
    "5m² Unit": 1.0,
    "10m² Unit": 1.5,
    "Premium Climate": 1.8
}

channels = ["Website", "Walk-in", "Google Ads", "Referral", "Partner"]


# HELPER FUNCTIONS
def seasonal_multiplier(month):
    # Summer demand higher
    if month in [5,6,7,8]:
        return 1.25
    elif month in [11,12,1,2]:
        return 0.85
    return 1.0

def random_duration():
    return random.randint(1, 12)

# GENERATE BOOKINGS
rows = []

for i in range(1, num_rows + 1):
    
    city = random.choice(list(cities.keys()))
    unit = random.choice(list(unit_sizes.keys()))
    
    start_date = fake.date_between(start_date="-2y", end_date="today")
    duration = random_duration()
    end_date = start_date + timedelta(days=duration*30)
    
    month = start_date.month
    
    base = cities[city]["base_price"]
    unit_mult = unit_sizes[unit]
    season_mult = seasonal_multiplier(month)
    
    price = round(base * unit_mult * season_mult, 2)
    
    repeat = random.choice([0,1])
    
    rows.append({
        "booking_id": i,
        "customer_id": random.randint(1000,9999),
        "city": city,
        "unit_size": unit,
        "start_date": start_date,
        "end_date": end_date,
        "duration_months": duration,
        "monthly_price": price,
        "total_revenue": round(price * duration,2),
        "channel": random.choice(channels),
        "repeat_customer": repeat
    })

df = pd.DataFrame(rows)

# Save CSV
df.to_csv("NordicSpace_Project/data/bookings.csv", index=False)

print("bookings.csv created successfully")
print(df.head())

# GENERATE REVIEWS

positive_reviews = [
    "Excellent service and clean unit.",
    "Easy booking process and friendly staff.",
    "Very secure location and fair pricing.",
    "Perfect solution during my move.",
    "Climate unit protected my furniture well."
]

negative_reviews = [
    "Late payment fee was unclear.",
    "Difficult to find the location first time.",
    "Customer service response was slow.",
    "Needed better signage outside.",
    "Price felt slightly high."
]

ratings = [2,3,4,5]

review_lines = []

for _ in range(30):
    city = random.choice(list(cities.keys()))
    month = random.choice(["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"])
    year = random.choice([2024,2025])
    
    if random.random() > 0.3:
        text = random.choice(positive_reviews)
        rating = random.choice([4,5])
    else:
        text = random.choice(negative_reviews)
        rating = random.choice([2,3])
    
    line = f'"{text}" - {city}, {month} {year}, {rating}/5'
    review_lines.append(line)

with open("NordicSpace_Project/data/reviews.txt", "w", encoding="utf-8") as f:
    for line in review_lines:
        f.write(line + "\n")

print("reviews.txt created successfully")



# Create folder structure
os.makedirs("data", exist_ok=True)

# Create workbook
wb = Workbook()
ws = wb.active
ws.title = "Targets"

# Add headers
ws.append(["KPI_Name", "Target_Value", "Unit", "Period"])

# Add rows manually
ws.append(["Occupancy Rate", 85, "%", "Monthly"])
ws.append(["Revenue per Unit", 120, "EUR", "Monthly"])
ws.append(["Avg Rental Duration", 8, "Months", "Yearly"])
ws.append(["Customer Satisfaction", 4.2, "/5.0", "Quarterly"])
ws.append(["Late Payment Rate", 5, "%", "Monthly"])

# Save file
wb.save("data/targets.xlsx")

print("targets.xlsx created successfully!")


import re

# Read raw review file
with open("data/reviews.txt", "r", encoding="utf-8") as f:
    lines = f.readlines()

rows = []

for line in lines:
    line = line.strip()

    # Example format:
    # "Excellent service." - Helsinki, Jan 2024, 5/5

    match = re.match(r'"(.*)" - ([A-Za-z]+), ([A-Za-z]{3}) (\d{4}), (\d)/5', line)

    if match:
        comment = match.group(1)
        city = match.group(2)
        month = match.group(3)
        year = match.group(4)
        rating = int(match.group(5))

        rows.append({
            "city": city,
            "month": month,
            "year": year,
            "rating": rating,
            "comment": comment
        })

feedback = pd.DataFrame(rows)

feedback.to_csv("data/feedback_cleaned.csv", index=False)

print("feedback_cleaned.csv created successfully")
feedback.head()

import requests


cities = {
    "Helsinki": (60.17, 24.94),
    "Espoo": (60.21, 24.66),
    "Vantaa": (60.29, 25.04),
    "Tampere": (61.50, 23.76),
    "Turku": (60.45, 22.27)
}

all_data = []

for city, (lat, lon) in cities.items():

    url = "https://archive-api.open-meteo.com/v1/archive"

    params = {
        "latitude": lat,
        "longitude": lon,
        "start_date": "2024-01-01",
        "end_date": "2025-12-31",
        "daily": "temperature_2m_mean",
        "timezone": "Europe/Helsinki"
    }

    response = requests.get(url, params=params)
    data = response.json()

    df = pd.DataFrame({
        "date": data["daily"]["time"],
        "avg_temp": data["daily"]["temperature_2m_mean"]
    })

    df["date"] = pd.to_datetime(df["date"])
    df["city"] = city
    df["year"] = df["date"].dt.year
    df["month"] = df["date"].dt.month

    monthly = df.groupby(["city", "year", "month"])["avg_temp"].mean().reset_index()

    all_data.append(monthly)

weather = pd.concat(all_data, ignore_index=True)

# Save in current folder first
weather.to_csv("data/weather_api.csv", index=False)

print("Final weather_api.csv created successfully")
weather.head(10)


import sqlite3

# Load files
bookings = pd.read_csv("data/bookings.csv")
weather = pd.read_csv("data/weather_api.csv")
feedback = pd.read_csv("data/feedback_cleaned.csv")

# Correct database location
conn = sqlite3.connect("database/nordicspace.db")

# Save tables
bookings.to_sql("bookings_raw", conn, if_exists="replace", index=False)
weather.to_sql("weather_raw", conn, if_exists="replace", index=False)
feedback.to_sql("feedback_raw", conn, if_exists="replace", index=False)

conn.close()

print("Database saved inside /database folder successfully.")



# Connect database
conn = sqlite3.connect("database/nordicspace.db")

# Load bookings
bookings = pd.read_sql("SELECT * FROM bookings_raw", conn)


# DIM CITY
dim_city = bookings[["city"]].drop_duplicates().reset_index(drop=True)
dim_city["city_id"] = dim_city.index + 1

# DIM UNIT
dim_unit = bookings[["unit_size", "monthly_price"]].drop_duplicates().reset_index(drop=True)
dim_unit["unit_id"] = dim_unit.index + 1


# FACT BOOKINGS

fact = bookings.merge(dim_city, on="city", how="left")
fact = fact.merge(dim_unit, on=["unit_size", "monthly_price"], how="left")

fact_bookings = fact[[
    "booking_id",
    "customer_id",
    "city_id",
    "unit_id",
    "duration_months",
    "monthly_price",
    "total_revenue",
    "repeat_customer"
]]

# Save tables
dim_city.to_sql("dim_city", conn, if_exists="replace", index=False)
dim_unit.to_sql("dim_unit", conn, if_exists="replace", index=False)
fact_bookings.to_sql("fact_bookings", conn, if_exists="replace", index=False)

conn.close()

print("Star schema created successfully.")
print("Tables:")
print("- dim_city")
print("- dim_unit")
print("- fact_bookings")

import shutil


# Backup database first
shutil.copy("database/nordicspace.db", "database/nordicspace_backup.db")

# Export SQL separately
conn = sqlite3.connect("database/nordicspace.db")

with open("database/nordicspace.sql", "w", encoding="utf-8") as f:
    for line in conn.iterdump():
        f.write(line + "\n")

conn.close()

print("Backup created: nordicspace_backup.db")
print("SQL exported: nordicspace.sql")

#### Project 2 & 3 ####
# Connect database
conn = sqlite3.connect("database/nordicspace.db")

# Read bookings data
df = pd.read_sql("SELECT * FROM bookings_raw", conn)

# KPI Calculations

total_revenue = df["total_revenue"].sum()
avg_duration = df["duration_months"].mean()
repeat_rate = df["repeat_customer"].mean() * 100

# Revenue by city
revenue_city = (
    df.groupby("city")["total_revenue"]
    .sum()
    .sort_values(ascending=False)
)

top_city = revenue_city.index[0]


# Print Results

print("DESCRIPTIVE ANALYTICS")
print("----------------------------")
print(f"Total Revenue: €{total_revenue:,.2f}")
print(f"Average Rental Duration: {avg_duration:.2f} months")
print(f"Repeat Customer Rate: {repeat_rate:.2f}%")
print(f"Top Performing City: {top_city}")

print("\nRevenue by City:")
print(revenue_city)

conn.close()



revenue_city.plot(kind="bar", figsize=(10,6))

plt.title("Revenue by City")
plt.xlabel("City")
plt.ylabel("Revenue (€)")
plt.xticks(rotation=0)
plt.tight_layout()

plt.savefig("charts/revenue_by_city.png")
plt.show()

from scipy.stats import linregress

# Connect database
conn = sqlite3.connect("database/nordicspace.db")

df = pd.read_sql("SELECT * FROM bookings_raw", conn)
conn.close()

# Convert date
df["start_date"] = pd.to_datetime(df["start_date"])

# Create monthly revenue
monthly = df.groupby(df["start_date"].dt.to_period("M"))["total_revenue"].sum().reset_index()
monthly["start_date"] = monthly["start_date"].astype(str)

# Add numeric index
monthly["month_num"] = range(1, len(monthly)+1)

# Regression model
x = monthly["month_num"]
y = monthly["total_revenue"]

slope, intercept, r, p, std_err = linregress(x, y)

# Predict next 3 months
future_months = [len(monthly)+1, len(monthly)+2, len(monthly)+3]
future_values = [intercept + slope*i for i in future_months]

print("Predicted Revenue Next 3 Months:")
for i, val in enumerate(future_values, 1):
    print(f"Month +{i}: €{val:,.2f}")

# Plot
plt.figure(figsize=(10,6))
plt.plot(monthly["month_num"], y, label="Actual Revenue")
plt.plot(future_months, future_values, "ro--", label="Forecast")

plt.title("Revenue Forecast (Next 3 Months)")
plt.xlabel("Month Number")
plt.ylabel("Revenue (€)")
plt.legend()
plt.tight_layout()

plt.savefig("charts/revenue_forecast.png")
plt.show()



conn = sqlite3.connect("database/nordicspace.db")
df = pd.read_sql("SELECT * FROM bookings_raw", conn)
conn.close()

# Revenue by city
city_rev = df.groupby("city")["total_revenue"].sum().sort_values(ascending=False)

best_city = city_rev.index[0]
worst_city = city_rev.index[-1]

repeat_rate = df["repeat_customer"].mean() * 100
avg_duration = df["duration_months"].mean()

print("PRESCRIPTIVE ANALYTICS RECOMMENDATIONS")
print("--------------------------------------")

print(f"1. Increase unit capacity in {best_city}, highest revenue location.")
print(f"2. Investigate performance improvement strategy for {worst_city}.")
print(f"3. Launch loyalty campaign (repeat rate = {repeat_rate:.1f}%).")
print(f"4. Promote long-term rental packages (avg stay = {avg_duration:.1f} months).")
print("5. Prepare seasonal pricing strategy for summer peak demand.")



conn = sqlite3.connect("database/nordicspace.db")
df = pd.read_sql("SELECT * FROM bookings_raw", conn)
conn.close()

df["start_date"] = pd.to_datetime(df["start_date"])
df["month"] = df["start_date"].dt.month

# KPI calculations
kpis = {
    "Total Revenue (€)": df["total_revenue"].sum(),
    "Total Bookings": len(df),
    "Average Rental Duration (Months)": df["duration_months"].mean(),
    "Repeat Customer Rate (%)": df["repeat_customer"].mean() * 100,
    "Average Price per Booking (€)": df["monthly_price"].mean(),
    "Top Performing City": df.groupby("city")["total_revenue"].sum().idxmax(),
    "Highest Revenue Month": df.groupby("month")["total_revenue"].sum().idxmax(),
    "Average Monthly Revenue (€)": df.groupby("month")["total_revenue"].sum().mean()
}

kpi_df = pd.DataFrame(list(kpis.items()), columns=["KPI_Name", "Value"])

kpi_df.to_excel("kpis.xlsx", index=False)

print("kpis.xlsx created successfully")
kpi_df



conn = sqlite3.connect("database/nordicspace.db")
df = pd.read_sql("SELECT * FROM bookings_raw", conn)
conn.close()

df["start_date"] = pd.to_datetime(df["start_date"])
df["month_year"] = df["start_date"].dt.to_period("M").astype(str)

monthly = df.groupby("month_year")["total_revenue"].sum()

plt.figure(figsize=(12,6))
monthly.plot(marker="o")

plt.title("Monthly Revenue Trend")
plt.xlabel("Month")
plt.ylabel("Revenue (€)")
plt.xticks(rotation=45)
plt.tight_layout()

plt.savefig("charts/monthly_revenue.png")
plt.show()


# Load data
conn = sqlite3.connect("database/nordicspace.db")
df = pd.read_sql("SELECT * FROM bookings_raw", conn)
conn.close()

df["start_date"] = pd.to_datetime(df["start_date"])
df["month_year"] = df["start_date"].dt.to_period("M").astype(str)


# 1 Revenue by City

plt.figure(figsize=(10,6))
df.groupby("city")["total_revenue"].sum().sort_values().plot(kind="bar")
plt.title("Revenue by City")
plt.xlabel("City")
plt.ylabel("Revenue (€)")
plt.tight_layout()
plt.savefig("charts/revenue_by_city.png")
plt.show()

# 2 Monthly Revenue Trend

plt.figure(figsize=(12,6))
df.groupby("month_year")["total_revenue"].sum().plot(marker="o")
plt.title("Monthly Revenue Trend")
plt.xlabel("Month")
plt.ylabel("Revenue (€)")
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig("charts/monthly_revenue.png")
plt.show()


# 3 Repeat Customer Pie

repeat_counts = df["repeat_customer"].value_counts()

plt.figure(figsize=(8,8))
plt.pie(repeat_counts,
        labels=["Repeat", "New"],
        autopct="%1.1f%%",
        startangle=90)
plt.title("Customer Type Distribution")
plt.tight_layout()
plt.savefig("charts/repeat_customer_pie.png")
plt.show()

# 4 Avg Price by City

plt.figure(figsize=(10,6))
df.groupby("city")["monthly_price"].mean().sort_values().plot(kind="bar")
plt.title("Average Price by City")
plt.xlabel("City")
plt.ylabel("Average Price (€)")
plt.tight_layout()
plt.savefig("charts/avg_price_by_city.png")
plt.show()


# 5 Bookings by City

plt.figure(figsize=(10,6))
df["city"].value_counts().plot(kind="bar")
plt.title("Bookings by City")
plt.xlabel("City")
plt.ylabel("Bookings")
plt.tight_layout()
plt.savefig("charts/bookings_by_city.png")
plt.show()


# 6 Rental Duration Histogram

plt.figure(figsize=(10,6))
plt.hist(df["duration_months"], bins=12)
plt.title("Rental Duration Distribution")
plt.xlabel("Months")
plt.ylabel("Frequency")
plt.tight_layout()
plt.savefig("charts/rental_duration_histogram.png")
plt.show()

print("All KPI charts created successfully.")

md = pd.DataFrame({
    "Category": [
        "Measure",
        "Measure",
        "Measure",
        "Measure",
        "Dimension",
        "Dimension",
        "Dimension",
        "Dimension",
        "Dimension"
    ],
    
    "Field_Name": [
        "total_revenue",
        "monthly_price",
        "duration_months",
        "booking_count",
        "city",
        "start_date",
        "month",
        "unit_size",
        "channel"
    ],
    
    "Description": [
        "Total sales revenue",
        "Monthly rental price",
        "Rental duration in months",
        "Number of bookings",
        "Location of booking",
        "Booking start date",
        "Month of booking",
        "Storage unit type",
        "Sales channel source"
    ]
})

md.to_excel("measurements_dimensions.xlsx", index=False)

print("measurements_dimensions.xlsx created successfully")
md


### project 4 ###

conn = sqlite3.connect("database/nordicspace.db")
df = pd.read_sql("SELECT * FROM bookings_raw", conn)
conn.close()

df["start_date"] = pd.to_datetime(df["start_date"])
df["month_year"] = df["start_date"].dt.to_period("M").astype(str)

monthly = df.groupby("month_year")["total_revenue"].sum().reset_index()

# Rolling average
monthly["rolling_avg"] = monthly["total_revenue"].rolling(3).mean()

plt.figure(figsize=(12,6))
plt.plot(monthly["month_year"], monthly["total_revenue"], label="Revenue")
plt.plot(monthly["month_year"], monthly["rolling_avg"], linewidth=3, label="3-Month Avg")

plt.title("Revenue Trend with Rolling Average")
plt.xlabel("Month")
plt.ylabel("Revenue (€)")
plt.xticks(rotation=45)
plt.legend()
plt.tight_layout()

plt.savefig("charts/trend_revenue_rolling.png")
plt.show()


import seaborn as sns


conn = sqlite3.connect("database/nordicspace.db")
df = pd.read_sql("SELECT * FROM bookings_raw", conn)
conn.close()

# Numeric columns only
corr_df = df[[
    "duration_months",
    "monthly_price",
    "total_revenue",
    "repeat_customer"
]]

corr = corr_df.corr()

plt.figure(figsize=(8,6))
sns.heatmap(corr, annot=True, cmap="coolwarm", fmt=".2f")

plt.title("Correlation Matrix")
plt.tight_layout()

plt.savefig("charts/correlation_heatmap.png")
plt.show()



benchmark = pd.DataFrame({
    "KPI": ["Repeat Rate %", "Avg Duration", "Occupancy Goal"],
    "NordicSpace": [49.6, 6.5, 85],
    "Industry Avg": [40, 5.5, 80]
})

benchmark.set_index("KPI").plot(kind="bar", figsize=(10,6))

plt.title("NordicSpace vs Industry Benchmark")
plt.ylabel("Value")
plt.xticks(rotation=0)
plt.tight_layout()

plt.savefig("charts/benchmark_comparison.png")
plt.show()