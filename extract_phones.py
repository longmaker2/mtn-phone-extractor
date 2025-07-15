import pandas as pd
import argparse
import os
from pathlib import Path


def extract_unique_phone_numbers_from_file(excel_file):
    """Extract unique phone numbers from a single Excel file."""
    try:
        # Try reading the Excel file with correct sheet name 'Sheet0'
        df = pd.read_excel(excel_file, sheet_name='Sheet0',
                           engine='xlrd' if excel_file.endswith('.xls') else 'openpyxl')
    except Exception as e:
        print(f"Error reading Excel file {excel_file}: {e}")
        return set()

    # Normalize column names by stripping whitespaces
    df.columns = [col.strip() for col in df.columns]

    # Extract phone numbers from column 'MSISDN'
    if 'MSISDN' not in df.columns:
        print(f"Column 'MSISDN' not found in Excel sheet: {excel_file}")
        return set()

    # Convert to string and filter out NaN values, then convert to set for uniqueness
    phone_numbers = df['MSISDN'].dropna().astype(str).tolist()

    # Remove any potential '.0' suffix from float conversion
    phone_numbers = [phone.replace('.0', '')
                     for phone in phone_numbers if phone != 'nan']

    print(
        f"📊 Extracted {len(phone_numbers)} phone numbers from {os.path.basename(excel_file)}")
    return set(phone_numbers)


def extract_unique_phone_numbers(input_path, output_file):
    """Extract unique phone numbers from Excel file(s)."""
    all_phone_numbers = set()
    total_raw_numbers = 0
    files_processed = 0

    if os.path.isfile(input_path):
        # Single file processing
        phone_numbers = extract_unique_phone_numbers_from_file(input_path)
        all_phone_numbers.update(phone_numbers)
        files_processed = 1
    elif os.path.isdir(input_path):
        # Directory processing - process all Excel files
        excel_files = os.scandir(input_path)
        if not excel_files:
            print(f"❌ No Excel files found in directory: {input_path}")
            return

        print(f"🔍 Found {len(list(excel_files))} Excel files to process")

        for excel_file in excel_files:
            if excel_file.name.endswith(('.xlsx', '.xls')):
                full_file_path = os.path.join(input_path, excel_file.name)
                print(full_file_path)
                phone_numbers = extract_unique_phone_numbers_from_file(full_file_path)
                files_processed += 1

            # Count raw numbers for deduplication tracking
            try:
                df = pd.read_excel(full_file_path, sheet_name='Sheet0',
                                   engine='xlrd' if full_file_path.endswith('.xls') else 'openpyxl')
                df.columns = [col.strip() for col in df.columns]
                if 'MSISDN' in df.columns:
                    raw_count = len(df['MSISDN'].dropna())
                    total_raw_numbers += raw_count
            except:
                pass  # Skip counting if file has issues

            all_phone_numbers.update(phone_numbers)
    else:
        print(f"❌ Input path does not exist: {input_path}")
        return

    if not all_phone_numbers:
        print("❌ No phone numbers extracted")
        return

    # Create new DataFrame and save to CSV
    # Sort the phone numbers for consistent output
    sorted_phone_numbers = sorted(list(all_phone_numbers))
    output_df = pd.DataFrame({'Phone Number': sorted_phone_numbers})

    # Verify no duplicates in final output
    if len(sorted_phone_numbers) != len(output_df['Phone Number'].unique()):
        print("⚠️  WARNING: Duplicates detected in final output!")
    else:
        print("✅ Verification: No duplicates in final output")

    output_df.to_csv(output_file, index=False)

    # Enhanced summary with deduplication stats
    duplicates_removed = total_raw_numbers - \
        len(sorted_phone_numbers) if total_raw_numbers > 0 else 0
    print(f"📊 DEDUPLICATION SUMMARY:")
    print(f"   Files processed: {files_processed}")
    if total_raw_numbers > 0:
        print(f"   Total raw entries: {total_raw_numbers:,}")
        print(f"   Duplicates removed: {duplicates_removed:,}")
        print(
            f"   Deduplication rate: {(duplicates_removed/total_raw_numbers)*100:.1f}%")
    print(f"✅ Final unique phone numbers: {len(sorted_phone_numbers):,}")
    print(f"✅ Saved to: {output_file}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description='Extract unique phone numbers from Excel file(s).')
    parser.add_argument(
        'input_path', help='Path to the Excel file or directory containing Excel files')
    parser.add_argument('--output', default='unique_phone_numbers.csv',
                        help='Output CSV file name (default: unique_phone_numbers.csv)')

    args = parser.parse_args()
    extract_unique_phone_numbers(args.input_path, args.output)
