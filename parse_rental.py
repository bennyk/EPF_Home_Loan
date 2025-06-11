from datetime import datetime
from dateutil.parser import parse
import re

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


# Function to parse date and check if within the last year
def is_within_last_year(date_str, current_date):
    try:
        post_date = parse(date_str)
        return (current_date - post_date).days <= 365
    except ValueError:
        return False


# Current date
current_date = datetime(2025, 6, 9)

# Filter listings within the last year and parse details
recent_listings = []
for listing in listings_data:
    if is_within_last_year(listing["date"], current_date):
        parsed_listing = {
            "agent": listing["agent"],
            "date": listing["date"],
            "price_rm": parse_price(listing["price"]),
            "price_per_sqft_rm": parse_price_per_sqft(listing["price_per_sqft"]),
            "size_sqft": parse_size(listing["size"]),
            "bedrooms": listing["bedrooms"],
            "bathrooms": listing["bathrooms"],
            "parking": listing["parking"],
            "furnished": listing["furnished"]
        }
        recent_listings.append(parsed_listing)

# Print parsed listings
print(f"Found {len(recent_listings)} rental listings within the last year:")
dates = []
prices = []
prices_per_sq = []

# Find prices and dates that fit perice per sqft
for listing in recent_listings:
    print("\nListing Details:")
    print(f"Agent: {listing['agent']}")
    print(f"Date: {listing['date']}")
    dates.append(listing['date'])
    print(f"Price: RM {listing['price_rm']:.2f}")
    prices.append(listing['price_rm'])
    print(f"Price per sq. ft.: RM {listing['price_per_sqft_rm']:.2f}")
    prices_per_sq.append(listing['price_per_sqft_rm'])
    print(f"Size: {listing['size_sqft']:.0f} sq. ft.")
    print(f"Bedrooms: {listing['bedrooms']}")
    print(f"Bathrooms: {listing['bathrooms']}")
    print(f"Parking: {listing['parking']}")
    print(f"Furnished: {listing['furnished']}")

print("\nAverage price per sqft is", sum(prices_per_sq)/len(prices_per_sq))
pass

