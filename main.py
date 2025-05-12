import math
import csv
from io import StringIO

# Loan and EPF parameters
initial_loan = 460000  # RM 460,000
interest_rate = 0.042  # 4.2% annual interest rate
tenure_months = 264  # 264 months (22 years)
epf_balance = 100000  # RM 100,000 in EPF Account 2 (example)
epf_dividend_rate = 0.055  # 5.5% average annual dividend rate (conservative)


# Function to calculate monthly loan payment
def calculate_monthly_payment(principal, annual_rate, months):
    monthly_rate = annual_rate / 12
    payment = principal * (monthly_rate * (1 + monthly_rate) ** months) / ((1 + monthly_rate) ** months - 1)
    return payment


# Function to generate amortization schedule
def generate_amortization_schedule(principal, annual_rate, months, output_file):
    monthly_rate = annual_rate / 12
    monthly_payment = calculate_monthly_payment(principal, annual_rate, months)
    balance = principal
    schedule = []

    for month in range(1, months + 1):
        interest = balance * monthly_rate
        principal_paid = monthly_payment - interest
        balance -= principal_paid
        if balance < 0:
            balance = 0
        schedule.append({
            'Month': month,
            'Payment': monthly_payment,
            'Interest': interest,
            'Principal': principal_paid,
            'Balance': balance
        })

    # Simulate writing to CSV
    output = StringIO()
    writer = csv.DictWriter(output, fieldnames=['Month', 'Payment', 'Interest', 'Principal', 'Balance'])
    writer.writeheader()
    for row in schedule:
        writer.writerow({
            'Month': row['Month'],
            'Payment': f"{row['Payment']:.2f}",
            'Interest': f"{row['Interest']:.2f}",
            'Principal': f"{row['Principal']:.2f}",
            'Balance': f"{row['Balance']:.2f}"
        })

    return schedule, output.getvalue()


# Function to print concise amortization schedule
def print_concise_schedule(schedule, title):
    print(f"\n{title}")
    print("Month | Payment | Interest | Principal | Balance")
    print("-" * 50)
    # Print first 3 months, last 3 months
    for row in schedule[:3] + schedule[-3:]:
        print(
            f"{row['Month']:>5} | RM {row['Payment']:>7.2f} | RM {row['Interest']:>7.2f} | RM {row['Principal']:>8.2f} | RM {row['Balance']:>9.2f}")


# Function to calculate EPF growth with compound interest
def calculate_epf_growth(principal, annual_rate, years):
    future_value = principal * (1 + annual_rate) ** years
    return future_value


# Scenario 1: Original loan without prepayment
monthly_payment_original = calculate_monthly_payment(initial_loan, interest_rate, tenure_months)
schedule_original, csv_original = generate_amortization_schedule(initial_loan, interest_rate, tenure_months,
                                                                 "original_loan_amortization.csv")
total_interest_original = sum(row['Interest'] for row in schedule_original)

# Scenario 2: Pay RM 100,000 from EPF to reduce loan
new_principal = initial_loan - epf_balance
monthly_payment_reduced = calculate_monthly_payment(new_principal, interest_rate, tenure_months)
schedule_reduced, csv_reduced = generate_amortization_schedule(new_principal, interest_rate, tenure_months,
                                                               "reduced_loan_amortization.csv")
total_interest_reduced = sum(row['Interest'] for row in schedule_reduced)
interest_saved = total_interest_original - total_interest_reduced

# EPF growth if RM 100,000 is kept invested
epf_future_value = calculate_epf_growth(epf_balance, epf_dividend_rate, tenure_months / 12)
epf_growth = epf_future_value - epf_balance

# Net financial impact
net_impact_payoff = interest_saved - epf_growth  # Negative means loss by paying off
net_impact_keep = epf_growth - total_interest_original  # Positive means gain by keeping

# Print results
print("=== Home Loan vs EPF Analysis with Amortization ===")
print("\nScenario 1: Continue loan without prepayment")
print(f"Monthly Payment: RM {monthly_payment_original:.2f}")
print(f"Total Interest Paid: RM {total_interest_original:.2f}")
print(f"Total Repayment: RM {monthly_payment_original * tenure_months:.2f}")
print_concise_schedule(schedule_original, "Amortization Schedule (Original Loan)")

print("\nScenario 2: Pay RM 100,000 from EPF to reduce loan")
print(f"New Principal: RM {new_principal:.2f}")
print(f"New Monthly Payment: RM {monthly_payment_reduced:.2f}")
print(f"Total Interest Paid: RM {total_interest_reduced:.2f}")
print(f"Total Repayment: RM {monthly_payment_reduced * tenure_months:.2f}")
print(f"Interest Saved: RM {interest_saved:.2f}")
print_concise_schedule(schedule_reduced, "Amortization Schedule (Reduced Loan)")

print("\nEPF Growth if RM 100,000 kept invested")
print(f"Future Value after 22 years: RM {epf_future_value:.2f}")
print(f"Growth (Profit): RM {epf_growth:.2f}")

print("\nFinancial Comparison")
print(f"Net Impact of Paying Off RM 100,000: RM {net_impact_payoff:.2f}")
print(f" (Negative means loss compared to keeping in EPF)")
print(f"Net Impact of Keeping RM 100,000 in EPF: RM {net_impact_keep:.2f}")
print(f" (Positive means gain compared to paying off loan)")

print("\nRecommendation:")
if net_impact_keep > 0:
    print("Keeping funds in EPF is likely better, as the dividend rate (5.5%) exceeds the loan interest rate (4.2%).")
else:
    print("Paying off the loan may be better, as the interest savings outweigh the EPF growth.")

# Output CSV content (simulated)
print("\nNote: Amortization schedules are generated as CSV content.")
print(
    "To save as files, copy the CSV content below to 'original_loan_amortization.csv' and 'reduced_loan_amortization.csv'.")
print("\n--- Original Loan Amortization CSV (First 5 rows) ---")
print("\n".join(csv_original.splitlines()[:6]))
print("\n--- Reduced Loan Amortization CSV (First 5 rows) ---")
print("\n".join(csv_reduced.splitlines()[:6]))