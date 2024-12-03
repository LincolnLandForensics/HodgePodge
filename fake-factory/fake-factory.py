#!/usr/bin/env python3

import os
import re
import sys
import random
import string
from faker import Faker
from argparse import ArgumentParser
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill

# Metadata
Author = "LincolnLandForensics"
version = "0.1.8"
description2 = 'Create fake data for research and testing. Never use live data in a test database!'
# Global color variables
color_red = "\033[31m"
color_yellow = "\033[33m"
color_green = "\033[32m"
color_reset = "\033[0m"

# Main function
def main():
    parser = ArgumentParser(description="Generate fake identities")
    parser.add_argument("-o", "--output", help="Output Excel file", required=False)
    parser.add_argument("-f", "--fakes", help="Generate fake IDs", action="store_true", required=False)
    parser.add_argument("-n", "--number", help="Number of records to generate", type=int, default=69)

    args = parser.parse_args()

    # global output_file
    if not args.output: 
        output_file = "FakeIdentities.xlsx"        
    else:
        output_file = args.output

    total = args.number
    total = check_number(total)

    if args.fakes:
        generate_fake_data(output_file, total)
    else:
        usage()


def check_number(number):
    if number > 1000000:
        print(f"{color_red}Number can't be larger than a million, I'm going to give you 50 for now.{color_reset}")
        number = 50  # Assign the default value
    return number

    
# Generate fake data
def generate_fake_data(output_file, total):
    # fake = Faker()
    fake = Faker('en_US')
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "FakeIdentities"

    # Freeze cell B2
    sheet.freeze_panes = "B2"  # Freeze the top row and the first column

    # Define headers
    # headers = [
        # "ip", "info", "fullname", "url", "email", "user", "phone", "business"
        # , "fulladdress", "city", "state", "zipcode", "note", "AKA", "DOB", "SEX"
        # , "SSN", "mothersmaidenname", "firstname", "middlename", "lastname"
    # ]

    headers = [
        "query", "ranking", "fullname", "url", "email", "user", "phone",
        "business", "fulladdress", "city", "state", "country", "zipcode", "AKA",
        "DOB", "SEX", "info", "mothers_maiden_name", "firstname", "middlename", "lastname",
        "associates", "case", "sosfilenumber", "owner", "president", "sosagent",
        "managers", "Time", "Latitude", "Longitude", "Coordinate",
        "original_file", "Source", "Source file information", "Plate", "VIS", "VIN",
        "VYR", "VMA", "LIC", "LIY", "DLN", "DLS", "content", "referer", "osurl",
        "titleurl", "pagestatus", "ip", "dnsdomain", "Tag", "Icon", "Type"
    ]


    # Write headers
    for col, header in enumerate(headers, start=1):
        cell = sheet.cell(row=1, column=col, value=header)
        fill = PatternFill(start_color="FFD700", end_color="FFD700", fill_type="solid")
        cell.fill = fill

    # Generate fake data rows
    for row in range(2, total + 2):
        ip = fake.ipv4()
        url = fake.url()
        fullname = fake.name()
        firstname, lastname = fullname.split()[:2]
        middle_initial = random.choice(string.ascii_uppercase)
        ssn = fake.ssn()
        # phone = fake.phone_number()
        phone = us_phone_number(fake)
        mothers_maiden = fake.last_name()
        gender = random.choice(["M", "F"])
        # email = fake.email()
        username = f"{firstname.lower()}.{lastname.lower()}"
        country = 'US'
        dob = generate_random_dob()
        city = fake.city()
        # state = fake.state()
        state = fake.state_abbr()
        zip_code = fake.postcode()
        fulladdress = f"{fake.street_address()}, {city}, {state} {zip_code}"
        business = fake.company()
        email = generate_random_email(firstname, lastname, business)
        plate = fake.license_plate()
        associates = fake.name()
        owner = fake.name()
        president = fake.name()
        sosagent = fake.name()
        managers = fake.name()
        VYR = get_random_year()
        VIN = generate_fake_vin()
        DLN = generate_fake_license_number()
        job = fake.job()
        uuid = fake.uuid4()
        # country = fake.country()
        # country = 'US'
        # language = fake.language_name()
        # profile = fake.simple_profile()
        # user_agent = fake.user_agent()
        # isbn = fake.isbn13()
        # currency = fake.currency_name()
 
        data = [
            "", "99 - fake data", fullname, url, email, username, phone, business, fulladdress, city, state, country, zip_code, "", dob, gender, ssn, mothers_maiden, firstname, middle_initial, lastname, associates, "", "", owner, president, sosagent, managers, "", "", "", "", "", "", "", plate, state, VIN, VYR, "", "", "", DLN, state, "", "", "", "", "", ip, "", "", "", ""
            ]

        for col, value in enumerate(data, start=1):
            sheet.cell(row=row, column=col, value=value)

    workbook.save(output_file)

    print(f'\n{color_green}{total} fake identities created. \nFake data saved to {output_file}{color_reset}\n') 
import random

def generate_fake_license_number():
    """
    Generates a fake driver's license number in the format T555-5555-5555.
    'T' is a random uppercase letter, and each '5' represents a digit.
    
    Returns:
        str: A fake driver's license number.
    """
    # Generate a random uppercase letter for the first character
    letter = random.choice(string.ascii_uppercase)
    
    # Generate the numeric sections
    numeric_section = '-'.join(
        ''.join(random.choices(string.digits, k=4)) for _ in range(3)
    )
    
    # Combine the parts
    license_number = f"{letter}{numeric_section}"
    return license_number
    
def generate_fake_vin():
    """
    Generates a fake VIN (Vehicle Identification Number).
    A VIN is a 17-character alphanumeric string, excluding 'I', 'O', and 'Q'.
    """
    # Define valid characters for a VIN
    vin_characters = string.ascii_uppercase + string.digits
    vin_characters = vin_characters.replace('I', '').replace('O', '').replace('Q', '')
    
    # Generate a 17-character string using valid VIN characters
    fake_vin = ''.join(random.choices(vin_characters, k=17))
    return fake_vin
    
    
def generate_random_email(firstname, lastname, business):
    '''
    Randomly generate a variety of email addresses based on firstname, lastname, 
    first_initial, and a random list of domains such as example.com, example.net, 
    email.com, email.net. Some emails will have random numbers at the end.

    :param firstname: First name of the person
    :param lastname: Last name of the person
    :return: Randomly generated email address
    '''

    fake = Faker()
    # List of possible email domains
    domains = [fake.free_email_domain() for _ in range(10)]

    # Replace spaces with commas, remove 'llc' and 'plc', and add '.net'
    business = business.replace(" ", "")  # Replace spaces
    business = business.replace(",", "")  # Replace commas
    business = business.replace("llc", "")  # Remove 'llc'
    business = business.replace("plc", "")  # Remove 'plc'
    business = business.strip()  # Strip any leading/trailing spaces

    # Append '.net' to the end of the business name
    business += ".net"

    # Append the modified business name to the domains list
    domains.append(business)

    # Randomly select a domain
    domain = random.choice(domains)

    # Randomly choose a format for the email
    email_formats = [
        f"{firstname.lower()}.{lastname.lower()}@{domain}",           # firstname.lastname
        f"{firstname[0].lower()}{lastname.lower()}@{domain}",         # first initial + lastname
        f"{firstname.lower()}{lastname.lower()}@{domain}",            # firstname + lastname
        f"{firstname[0].lower()}{lastname.lower()}{random.randint(1, 99)}@{domain}",  # first initial + lastname + number
        f"{firstname.lower()}{random.randint(1, 99)}@{domain}",       # firstname + number
        f"{lastname.lower()}{random.randint(1, 99)}@{domain}"         # lastname + number
    ]

    # Select a random email format
    email = random.choice(email_formats)

    return email


def generate_random_dob():
    year = random.randint(1950, 2000)
    month = random.randint(1, 12)
    day = random.randint(1, 28)  # To avoid invalid dates
    return f"{year}-{month:02d}-{day:02d}"


def get_random_year():
    """Returns a random 4-digit year between 2005 and 2024."""
    return random.randint(2005, 2024)

    
def us_phone_number(fake):
    phone_number = fake.phone_number()
    # Ensure the phone number is in the format xxx-xxx-xxxx
    formatted_phone_number = re.sub(r'\D', '', phone_number)[:10]
    return f"{formatted_phone_number[:3]}-{formatted_phone_number[3:6]}-{formatted_phone_number[6:]}"
    
# Usage instructions
def usage():
    file = sys.argv[0].split('\\')[-1]
    print(f'\nDescription: {color_green}{description2}{color_reset}')
    print(f'{file} Version: {version} by {Author}')
    print(f"{color_yellow}Usage:{color_reset}")
    print(f"\tpython {sys.argv[0]} -f -n <number_of_records> -o <output_file>")
    print(f"{color_yellow}Example:{color_reset}")
    print(f"\tpython {sys.argv[0]} -f")
    print(f"\tpython {sys.argv[0]} -f -n 100 -o fake_identities.xlsx")

# Run the script
if __name__ == "__main__":
    main()
