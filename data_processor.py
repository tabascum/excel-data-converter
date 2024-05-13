import pandas as pd
import os
import re
import random
from datetime import datetime


def distribute_qty(df):
    df['Customer Name'] = df['Customer Name'].fillna('')

    saturn_rows = df[df['Customer Name'].str.contains('MEDIAMARKT SATURN')]
    for index, row in saturn_rows.iterrows():
        if pd.notna(row['Model']) and pd.notna(row['Promotion Name']):
            matching_rows = df[(df['Model'] == row['Model']) & 
                               (df['Promotion Name'] == row['Promotion Name']) &
                               (~df['Customer Name'].str.contains('MEDIAMARKT SATURN'))]

            if matching_rows.empty:
                continue

            total_qty = row['Expected QTY(Editable)']

            if total_qty > 11:
                distribute_evenly(df, total_qty, matching_rows, index)
            elif total_qty < 0:
                remove_randomly(df, abs(total_qty), matching_rows, index)
            else:  # total_qty <= 11 and total_qty >= 0
                distribute_randomly(df, total_qty, matching_rows, index)

    # Remove MEDIAMARKT SATURN rows and rows with Expected QTY(Editable) of 0
    df = df[~df['Customer Name'].str.contains('MEDIAMARKT SATURN')]
    df = df[df['Expected QTY(Editable)'] != 0]

    return df


def distribute_evenly(df, total_qty, matching_rows, index):
    num_customers = len(matching_rows)
    qty_per_customer = total_qty // num_customers
    remainder = total_qty % num_customers

    df.loc[matching_rows.index, 'Expected QTY(Editable)'] += qty_per_customer

    if remainder > 0:
        random_customer = random.choice(matching_rows.index)
        df.at[random_customer, 'Expected QTY(Editable)'] += remainder

    df.at[index, 'Expected QTY(Editable)'] = 0


def distribute_randomly(df, total_qty, matching_rows, index):
    for _ in range(total_qty):
        random_customer = random.choice(matching_rows.index)
        df.at[random_customer, 'Expected QTY(Editable)'] += 1

    df.at[index, 'Expected QTY(Editable)'] = 0


def remove_randomly(df, total_qty_to_remove, matching_rows, saturn_index):
    while total_qty_to_remove > 0:
        # Filter matching rows with Expected QTY(Editable) > 0
        eligible_rows = matching_rows[matching_rows['Expected QTY(Editable)'] > 0]
        if eligible_rows.empty:
            break  # Exit if no eligible rows to avoid infinite loop

        random_customer_index = random.choice(eligible_rows.index)
        current_qty = df.at[random_customer_index, 'Expected QTY(Editable)']

        if current_qty > 0:
            reduction = min(current_qty, total_qty_to_remove)
            df.at[random_customer_index, 'Expected QTY(Editable)'] -= reduction
            total_qty_to_remove -= reduction

    # Set MEDIAMARKT SATURN quantity to 0, assuming its negative value is now offset
    df.at[saturn_index, 'Expected QTY(Editable)'] = 0

def consolidate_duplicate_models(df):
    # Convert 'Apply Date(From)' and 'Apply Date(To)' to datetime objects
    df['Apply Date(From)'] = pd.to_datetime(df['Apply Date(From)'], format='%Y%m%d', errors='coerce')
    df['Apply Date(To)'] = pd.to_datetime(df['Apply Date(To)'], format='%Y%m%d', errors='coerce')

    # Create temporary year columns
    df['Year_From'] = df['Apply Date(From)'].dt.year
    df['Year_To'] = df['Apply Date(To)'].dt.year

    # Filter out MEDIAMARKT promotions
    non_mediamarkt_df = df[~df['Customer Name'].str.contains('MEDIAMARKT', na=False)]

    # Group by Model, Promotion Name, and year of Apply Date(From) and Apply Date(To)
    grouped = non_mediamarkt_df.groupby(['Model', 'Promotion Name', 'Year_From', 'Year_To'])

    for _, group in grouped:
        if len(group) > 1:
            # Sum the quantities
            total_qty = group['Expected QTY(Editable)'].sum()

            # Get index of the first row in the group
            first_index = group.index[0]

            # Update the first row with the total quantity
            df.at[first_index, 'Expected QTY(Editable)'] = total_qty

            # Drop the other rows in the group
            df = df.drop(group.index[1:])

    # Convert 'Apply Date(From)' and 'Apply Date(To)' back to the format YYYYMMDD
    df['Apply Date(From)'] = df['Apply Date(From)'].dt.strftime('%Y%m%d')
    df['Apply Date(To)'] = df['Apply Date(To)'].dt.strftime('%Y%m%d')

    # Remove temporary year columns
    df.drop(columns=['Year_From', 'Year_To'], inplace=True)

    return df


def process_excel(file_path, download_directory):
    df = pd.read_excel(file_path)

    df['Apply Date(From)'] = pd.to_datetime(df['Apply Date(From)'].astype(str), format='%Y%m%d', errors='coerce')
    df['Apply Date(To)'] = pd.to_datetime(df['Apply Date(To)'].astype(str), format='%Y%m%d', errors='coerce')

    current_year = datetime.now().year
    current_month_year = datetime.now().strftime("%Y%m")


    def transform_promotion_name(promotion_name, customer_name):
        # Convert both promotion_name and customer_name to strings
        promotion_name = str(promotion_name)
        customer_name = str(customer_name)

        # Extract the program number
        promotion_number = re.search(r'\b[A-Za-z]{2}\d{4,5}\b', promotion_name)
        promotion_number = promotion_number.group() if promotion_number else ""

        # Check if the customer is MEDIAMARKT
        if "MEDIAMARKT" in customer_name:
            # Replace 'MEDIAMARKT' with 'MM'
            customer_name = customer_name.replace('MEDIAMARKT', 'MM')
            
            # Extract any prefix from promotion_name
            prefix_match = re.search(r'^.*?(?=\sMM|\sMEDIAMARKT)', promotion_name)
            prefix = prefix_match.group() if prefix_match else promotion_name.rstrip()

            # Extract the customer suffix after 'MEDIAMARKT'
            parts = customer_name.split('MM')
            customer_suffix = parts[-1].strip() if len(parts) > 1 else ""

            # Construct the new promotion name
            new_promotion_name = f"{prefix} MM {customer_suffix}".strip()
            if promotion_number:
                return f"{new_promotion_name} - {promotion_number} - NP - E"
            else:
                return f"{new_promotion_name} - NP - E"
        else:
            return f"{promotion_name} - NP - E"


    df['Sales PGM Name(Editable)'] = df.apply(lambda row: transform_promotion_name(row['Promotion Name'], row['Customer Name']), axis=1)
    df['Registration Requeste Date(Editable)'] = datetime.now().strftime("%Y%m%d")
    df['Accounting Unit(Editable)'] = 'SAL'
    df['Department(Editable)'] = 20066

    df['Apply Month(Editable)'] = df['Apply Date(To)'].dt.strftime('%Y%m')

    df = distribute_qty(df)
    df = consolidate_duplicate_models(df)


    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    processed_file_name = f"processed_file_{timestamp}.xlsx"
    processed_file_path = os.path.join(download_directory, processed_file_name)
    df.to_excel(processed_file_path, index=False)
    return processed_file_path

    


