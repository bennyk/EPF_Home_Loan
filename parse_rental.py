from datetime import datetime
from dateutil.parser import parse
import re
import requests
from bs4 import BeautifulSoup
import statistics
import numpy as np

# TODO hard coded numbers
# Sample data from the document (simulating the parsed content)
listings_data = [
    {"agent": "Grace Chong", "date": "06 Jun 2025 03:45 AM", "price": "RM 4,100", "price_per_sqft": "RM 3.2 per sq. ft.", "size": "1,281 sq. ft.", "bedrooms": "3", "bathrooms": "2", "parking": "1", "furnished": "Not specified"},
    {"agent": "Jerry Lee", "date": "06 Jun 2025 03:01 AM", "price": "RM 4,800", "price_per_sqft": "RM 3.03 per sq. ft.", "size": "1,582 sq. ft.", "bedrooms": "3+1", "bathrooms": "3", "parking": "2", "furnished": "Fully furnished"},
    {"agent": "XiaXun Ong", "date": "05 Jun 2025 08:17 AM", "price": "RM 2,700", "price_per_sqft": "RM 3.6 per sq. ft.", "size": "750 sq. ft.", "bedrooms": "1", "bathrooms": "1", "parking": "1", "furnished": "Fully furnished"},
    {"agent": "Vickie Foong", "date": "04 Jun 2025 12:50 PM", "price": "RM 3,300", "price_per_sqft": "RM 2.58 per sq. ft.", "size": "1,281 sq. ft.", "bedrooms": "3", "bathrooms": "2", "parking": "2", "furnished": "Fully furnished"},
    {"agent": "Brian E", "date": "26 May 2025 11:56 PM", "price": "RM 2,400", "price_per_sqft": "RM 3.23 per sq. ft.", "size": "743 sq. ft.", "bedrooms": "1", "bathrooms": "1", "parking": "1", "furnished": "Fully furnished"},
    {"agent": "Brian E", "date": "14 May 2025 02:26 PM", "price": "RM 3,600", "price_per_sqft": "RM 2.81 per sq. ft.", "size": "1,281 sq. ft.", "bedrooms": "3", "bathrooms": "2", "parking": "1", "furnished": "Fully furnished"},
    {"agent": "Mei Low", "date": "14 May 2025 09:53 AM", "price": "RM 5,000", "price_per_sqft": "RM 3.44 per sq. ft.", "size": "1,453 sq. ft.", "bedrooms": "3", "bathrooms": "3", "parking": "1", "furnished": "Fully furnished"},
    {"agent": "Dan Lo", "date": "09 Jun 2025 03:59 AM", "price": "RM 4,300", "price_per_sqft": "RM 3.35 per sq. ft.", "size": "1,285 sq. ft.", "bedrooms": "3", "bathrooms": "2", "parking": "1", "furnished": "Fully furnished"},
    {"agent": "Calvin Poh", "date": "06 Jun 2025 02:40 AM", "price": "RM 2,800", "price_per_sqft": "RM 4.64 per sq. ft.", "size": "603 sq. ft.", "bedrooms": "1", "bathrooms": "1", "parking": "1", "furnished": "Fully furnished"},
    {"agent": "James Leong", "date": "05 Jun 2025 01:53 PM", "price": "RM 2,600", "price_per_sqft": "RM 3.47 per sq. ft.", "size": "750 sq. ft.", "bedrooms": "1", "bathrooms": "1", "parking": "1", "furnished": "Fully furnished"},
    {"agent": "Ann Ong", "date": "05 Jun 2025 01:45 AM", "price": "RM 3,800", "price_per_sqft": "RM 2.97 per sq. ft.", "size": "1,281 sq. ft.", "bedrooms": "3", "bathrooms": "2", "parking": "1", "furnished": "Not specified"},
    {"agent": "Brian E", "date": "04 Jun 2025 11:33 PM", "price": "RM 2,500", "price_per_sqft": "RM 3.36 per sq. ft.", "size": "743 sq. ft.", "bedrooms": "1", "bathrooms": "1", "parking": "1", "furnished": "Not specified"},
    {"agent": "James Leong", "date": "04 Jun 2025 03:19 PM", "price": "RM 2,699", "price_per_sqft": "RM 3.6 per sq. ft.", "size": "750 sq. ft.", "bedrooms": "1", "bathrooms": "1", "parking": "1", "furnished": "Fully furnished"},
    {"agent": "Christina Low", "date": "04 Jun 2025 04:35 AM", "price": "RM 4,500", "price_per_sqft": "RM 3.1 per sq. ft.", "size": "1,453 sq. ft.", "bedrooms": "3", "bathrooms": "3", "parking": "2", "furnished": "Fully furnished"},
    {"agent": "James Leong", "date": "04 Jun 2025 03:16 AM", "price": "RM 2,699", "price_per_sqft": "RM 3.6 per sq. ft.", "size": "750 sq. ft.", "bedrooms": "1", "bathrooms": "1", "parking": "1", "furnished": "Fully furnished"},
    {"agent": "Grace Chong", "date": "03 Jun 2025 07:36 AM", "price": "RM 4,300", "price_per_sqft": "RM 3.36 per sq. ft.", "size": "1,281 sq. ft.", "bedrooms": "3", "bathrooms": "2", "parking": "1", "furnished": "Not specified"},
    {"agent": "James Leong", "date": "01 Jun 2025 03:32 PM", "price": "RM 3,699", "price_per_sqft": "RM 2.89 per sq. ft.", "size": "1,281 sq. ft.", "bedrooms": "3", "bathrooms": "2", "parking": "1", "furnished": "Fully furnished"},
    {"agent": "Nck Chan", "date": "30 May 2025 06:37 PM", "price": "RM 5,000", "price_per_sqft": "RM 3.57 per sq. ft.", "size": "1,400 sq. ft.", "bedrooms": "3", "bathrooms": "3", "parking": "1", "furnished": "Fully furnished"},
    {"agent": "Ann Ong", "date": "30 May 2025 12:37 PM", "price": "RM 2,400", "price_per_sqft": "RM 3.28 per sq. ft.", "size": "732 sq. ft.", "bedrooms": "1", "bathrooms": "1", "parking": "1", "furnished": "Not specified"},
    {"agent": "Ann Ong", "date": "30 May 2025 12:34 PM", "price": "RM 2,500", "price_per_sqft": "RM 3.42 per sq. ft.", "size": "732 sq. ft.", "bedrooms": "1", "bathrooms": "1", "parking": "1", "furnished": "Not specified"},
]


# Function to parse RM price and convert to float
def parse_price(price_str):
    return float(re.sub(r'[^\d.]', '', price_str))


# Function to parse size and convert to float
def parse_size(size_str):
    return float(re.match(r'(\d+)', re.sub(r'[,]', '', size_str)).group(1))


# Function to parse price per square foot
def parse_price_per_sqft(price_per_sqft_str):
    m = re.match(r'[^\d.]+([\d.]+)', price_per_sqft_str)
    return float(m.group(1))


# Function to check if date is within the last year
def is_within_last_year(date_str, current_date):
    try:
        post_date = parse(date_str)
        return (current_date - post_date).days <= 365
    except ValueError:
        return False


# Function to scrape listings (placeholder)
def scrape_listings(url):
    try:
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/91.0.4472.124'}
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, 'html.parser')
        # TODO: Implement parsing based on website HTML
        print("Scraping not implemented; using sample data.")
        return listings_data
    except Exception as e:
        print(f"Error scraping: {e}. Using sample data.")
        return listings_data


# Phase 1: Determine dynamic price per sq. ft. range
def phase1_determine_range(listings, current_date):
    price_per_sqft_values = []

    for listing in listings:
        if is_within_last_year(listing["date"], current_date):
            price_per_sqft = parse_price_per_sqft(listing["price_per_sqft"])
            price_per_sqft_values.append(price_per_sqft)

    if not price_per_sqft_values:
        return None, None

    # Calculate Q1, Q3, and IQR
    q1 = np.percentile(price_per_sqft_values, 25)
    q3 = np.percentile(price_per_sqft_values, 75)
    iqr = q3 - q1

    # Determine range, capped at data's min and max
    min_price_per_sqft = max(min(price_per_sqft_values), q1 - 1.5 * iqr)
    min_price_per_sqft *= 1.2
    max_price_per_sqft = min(max(price_per_sqft_values), q3 + 1.5 * iqr)
    max_price_per_sqft *= .8

    return min_price_per_sqft, max_price_per_sqft


# Phase 2: Filter listings and compute statistics
def phase2_filter_and_stats(listings, min_price_per_sqft, max_price_per_sqft, current_date):
    matching_listings = []
    price_per_sqft_values = []

    for listing in listings:
        if is_within_last_year(listing["date"], current_date):
            price_per_sqft = parse_price_per_sqft(listing["price_per_sqft"])
            if min_price_per_sqft <= price_per_sqft <= max_price_per_sqft:
                parsed_listing = {
                    "agent": listing["agent"],
                    "date": listing["date"],
                    "price_rm": parse_price(listing["price"]),
                    "price_per_sqft_rm": price_per_sqft,
                    "size_sqft": parse_size(listing["size"]),
                    "bedrooms": listing["bedrooms"],
                    "bathrooms": listing["bathrooms"],
                    "parking": listing["parking"],
                    "furnished": listing["furnished"]
                }
                matching_listings.append(parsed_listing)
                price_per_sqft_values.append(price_per_sqft)

    # Calculate statistics
    stats = {}
    if price_per_sqft_values:
        stats["min"] = min(price_per_sqft_values)
        stats["max"] = max(price_per_sqft_values)
        stats["median"] = statistics.median(price_per_sqft_values)

    return matching_listings, stats


# Main script
url = "https://www.iproperty.com.my/rent/kl-city-centre/suasana-bukit-ceylon-raja-chulan-residences-y8wzqe/serviced-residence/?l1"
current_date = datetime(2025, 6, 11, 17, 8)  # 05:08 PM, June 11, 2025

# Get listings
listings = scrape_listings(url)

# Phase 1: Determine range
min_price_per_sqft, max_price_per_sqft = phase1_determine_range(listings, current_date)

if min_price_per_sqft is None or max_price_per_sqft is None:
    print("No valid listings found within the last year.")
    exit()

print(f"\nDynamically determined price per sq. ft. range: RM {min_price_per_sqft:.2f} to RM {max_price_per_sqft:.2f}")

# Phase 2: Filter and compute statistics
matching_listings, stats = phase2_filter_and_stats(listings, min_price_per_sqft, max_price_per_sqft, current_date)

# Print statistics
print(f"\nPrice per Sq. Ft. Statistics for Filtered Listings:")
if stats:
    print(f"Minimum: RM {stats['min']:.2f} per sq. ft.")
    print(f"Maximum: RM {stats['max']:.2f} per sq. ft.")
    print(f"Median: RM {stats['median']:.2f} per sq. ft.")
else:
    print("No listings match the dynamic range; no statistics available.")

# Print matching listings
print(f"\nFound {len(matching_listings)} listings within the dynamic price per sq. ft. range:")
if matching_listings:
    for listing in matching_listings:
        print("\nListing Details:")
        print(f"Agent: {listing['agent']}")
        print(f"Date: {listing['date']}")
        print(f"Price: RM {listing['price_rm']:.2f}")
        print(f"Price per sq. ft.: RM {listing['price_per_sqft_rm']:.2f}")
        print(f"Size: {listing['size_sqft']:.0f} sq. ft.")
        print(f"Bedrooms: {listing['bedrooms']}")
        print(f"Bathrooms: {listing['bathrooms']}")
        print(f"Parking: {listing['parking']}")
        print(f"Furnished: {listing['furnished']}")
else:
    print("No listings match the dynamic range.")