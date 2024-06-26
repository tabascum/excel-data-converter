import pandas as pd
import os
import re
import random
from datetime import datetime

def distribute_qty(df):
    df['Customer Name'] = df['Customer Name'].fillna('')

    saturn_rows = df[df['Customer Name'].str.contains('MEDIAMARKT SATURN')]
    for index, row in saturn_rows.iterrows():
        if pd.notna(row['Model(Editable)']) and pd.notna(row['Promotion Name']):
            matching_rows = df[(df['Model(Editable)'] == row['Model(Editable)']) & 
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
    # non_mediamarkt_df = df[~df['Customer Name'].str.contains('MEDIAMARKT', na=False)]

    # Group by Model, Promotion Name, Year_From, Year_To, and Amount Per Unit
    grouped = df.groupby(['Model(Editable)', 'Promotion Name', 'Year_From', 'Year_To', 'Amount Per Unit'])

    for _, group in grouped:
        if len(group) > 1:
            # Check if all Amount Per Unit values are the same in the group
            if group['Amount Per Unit'].nunique() == 1:
                # Sum the quantities
                total_qty = group['Expected QTY(Editable)'].sum()

                # Get index of the first row in the group
                first_index = group.index[0]

                # Update the first row with the total quantity
                df.at[first_index, 'Expected QTY(Editable)'] = total_qty

                # Drop the other rows in the group
                df = df.drop(group.index[1:])
            else:
                # If Amount Per Unit values are different, keep the duplicates
                pass

    # Convert 'Apply Date(From)' and 'Apply Date(To)' back to the format YYYYMMDD
    df['Apply Date(From)'] = df['Apply Date(From)'].dt.strftime('%Y%m%d')
    df['Apply Date(To)'] = df['Apply Date(To)'].dt.strftime('%Y%m%d')

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

        # Extract the prefix from the promotion name (e.g., "BG SO MM LEIRIA - Z02")
        prefix_match = re.search(r'^(.*?)\s*- Z\d+', promotion_name)
        prefix = prefix_match.group(1).strip() if prefix_match else promotion_name.strip()

        # Randomly select a customer name from MEDIAMARKT or MEDIA MARKT promotions
        mediamarkt_customers = re.findall(r'\bMEDI[ ]?A?[ ]?MARKT\s\S+', customer_name)  # Match MEDIAMARKT or MEDIA MARKT
        selected_customer_name = random.choice(mediamarkt_customers) if mediamarkt_customers else ''

        # Replace 'MEDIAMARKT' or 'MEDIA MARKT' with 'MM' and append customer suffix if applicable
        if "MEDIAMARKT" in promotion_name or "MEDIA MARKT" in promotion_name:
            customer_suffix = selected_customer_name.replace('MEDIAMARKT', '').replace('MEDIA MARKT', '').strip()
            # Remove any trailing comma after 'MEDIAMARKT' or 'MEDIA MARKT' in promotion_name
            promotion_name = re.sub(r'(MEDIAMARKT|MEDIA MARKT),?', 'MM', promotion_name, flags=re.IGNORECASE).strip()
            new_promotion_name = promotion_name.replace('MM', f'MM {customer_suffix}').strip()
        else:
            new_promotion_name = promotion_name

        # Check if the promotion name contains "UPDATE" or "RECREATE"
        if "UPDATE" in new_promotion_name or "RECREATE" in new_promotion_name:
            # Remove "UPDATE" or "RECREATE" from the promotion name
            new_promotion_name = re.sub(r'\s*- UPDATE\s*|\s*- RECREATE\s*', '', new_promotion_name, flags=re.IGNORECASE)

        return f"{new_promotion_name} - NP - E"



    df['Sales PGM Name(Editable)'] = df.apply(lambda row: transform_promotion_name(row['Promotion Name'], row['Customer Name']), axis=1)
    df['Registration Request Date(Editable)'] = datetime.now().strftime("%Y%m%d")
    df['Accounting Unit(Editable)'] = 'SAL'
    df['Department(Editable)'] = 20066
    df['Apply Month(Editable)'] = "202406"

    df = distribute_qty(df)
    df = consolidate_duplicate_models(df)

    # Calculate 'Expected Cost(Editable)'
    df['Expected Cost(Editable)'] = df['Amount Per Unit'] * df['Expected QTY(Editable)']

    # Maintain the original column order
    original_columns = pd.read_excel(file_path, nrows=0).columns.tolist()
    df = df[original_columns]

    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    processed_file_name = f"processed_file_{timestamp}.xlsx"
    processed_file_path = os.path.join(download_directory, processed_file_name)
    df.to_excel(processed_file_path, index=False)
    return processed_file_path