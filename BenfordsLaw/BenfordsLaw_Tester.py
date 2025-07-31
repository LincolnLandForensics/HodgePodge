#!/usr/bin/env python3
# coding: utf-8
"""
Read XLSX file analyze numbers with Benford's Law.
Author: LincolnLandForensics
Version: 0.1.2
"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Imports        >>>>>>>>>>>>>>>>>>>>>>>>>>

import os
import sys
import argparse

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from scipy.stats import chisquare



# Color support for Windows 11+
color_red = color_yellow = color_green = color_blue = color_purple = color_reset = ''
if sys.version_info > (3, 7, 9) and os.name == "nt":
    from colorama import Fore, Back, Style
    print(Back.BLACK)
    color_red, color_yellow, color_green = Fore.RED, Fore.YELLOW, Fore.GREEN
    color_blue, color_purple, color_reset = Fore.BLUE, Fore.MAGENTA, Style.RESET_ALL

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Pre-Sets       >>>>>>>>>>>>>>>>>>>>>>>>>>

author = 'LincolnLandForensics'
description = "This script models Benford’s Law by generating and comparing authentic versus manipulated financial data. It outputs frequency distributions to Excel for forensic analysis, helping identify statistical anomalies suggestive of fraud."
version = '0.0.1'


# <<<<<<<<<<<<<<<<<<<<<<<<<<      Menu           >>>>>>>>>>>>>>>>>>>>>>>>>>



def main():
    global row
    row = 0  # defines arguments
    # Row = 1  # defines arguments   # if you want to add headers 
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument('-I', '--input', help='', required=False)
    # parser.add_argument('-O', '--output', help='', required=False)
    parser.add_argument('-b', '--benfords', help='read xlsx', required=False, action='store_true')
    # parser.add_argument('-c', '--column', help='choose column', required=False, action='store_true')
    parser.add_argument('-c', '--column', help='', required=False)
    args = parser.parse_args()


    global input_file
    input_file = args.input if args.input else "benfordsLaw_tester.xlsx"

    global column_pick
    column_pick = args.column if args.column else "A"


    if args.benfords:

        file_exists = os.path.exists(input_file)
        if file_exists == True:
            msg_blurb = (f'Reading {input_file} column {column_pick}')
            msg_blurb_square(msg_blurb, color_green)    
            
            # data = read_xlsx(input_file)
            benfords(input_file)

        else:
            msg_blurb = (f'{input_file} does not exist')
            msg_blurb_square(msg_blurb, color_red)      
            exit()

    else:
        usage()
    
    return 0

def benfords(input_file):

    # 👀 Load one column from Excel using openpyxl
    wb = load_workbook('benfordsLaw_tester.xlsx', data_only=True)
    ws = wb.active
    print(f'{wb.active}')   # temp
    column_data = [cell.value for cell in ws[column_pick] if isinstance(cell.value, (int, float))]

    # 📊 Generate synthetic data for Benford analysis
    np.random.seed(123)
    authentic_revenue = 10 ** np.random.uniform(np.log10(1000), np.log10(500000), 1000)
    fraudulent_revenue = np.round(np.random.uniform(1000, 500000, 1000), -3)

    # 🔍 First-digit extraction
    def get_first_digit(arr):
        return [int(str(int(np.floor(x)))[0]) for x in arr if x > 0]

    authentic_first_digit = get_first_digit(authentic_revenue)
    fraudulent_first_digit = get_first_digit(fraudulent_revenue)
    excel_first_digit = get_first_digit(column_data)

    # 📈 Benford expected distribution
    benford_dist = pd.DataFrame({
        'Digit': range(1, 10),
        'Expected': np.log10(1 + 1 / np.arange(1, 10))
    })

    # 🧮 Convert digit list to DataFrame
    def digit_df(digits, label):
        df = pd.Series(digits).value_counts().sort_index().reset_index()
        df.columns = ['Digit', 'Freq']
        df['Prop'] = df['Freq'] / df['Freq'].sum()
        df['Type'] = label
        return pd.merge(df, benford_dist, on='Digit', how='left')

    authentic_df = digit_df(authentic_first_digit, 'Authentic Data')
    fraudulent_df = digit_df(fraudulent_first_digit, 'Manipulated Data')
    excel_df = digit_df(excel_first_digit, 'Excel Data')

    # 📊 Combine datasets for plotting
    data_plot = pd.concat([authentic_df, fraudulent_df, excel_df])

    # 📐 Chi-Square Test (Excel vs Benford)
    expected_counts = benford_dist['Expected'] * len(excel_first_digit)
    observed_counts = excel_df['Freq']
    chi2_stat, p_value = chisquare(f_obs=observed_counts, f_exp=expected_counts)
    print(f"Chi-Square Test ➤ Stat: {chi2_stat:.2f}, p-value: {p_value:.4f}")

    # 📏 MAD Test (Excel vs Benford)
    mad = np.mean(np.abs(excel_df['Prop'] - excel_df['Expected']))
    print(f"MAD ➤ Mean Absolute Deviation: {mad:.4f}")

    # 📉 Plotting
    fig, ax = plt.subplots(figsize=(10, 6))
    colors = {
        'Authentic Data': 'darkgreen',
        'Manipulated Data': 'red',
        'Excel Data': 'orange'
    }
    offsets = {
        'Authentic Data': -0.25,
        'Manipulated Data': 0.25,
        'Excel Data': 0.0
    }
    for label, group in data_plot.groupby('Type'):
        ax.bar(group['Digit'] + offsets[label],
               group['Prop'], width=0.25, alpha=0.7,
               label=label, color=colors[label])

    ax.plot(benford_dist['Digit'], benford_dist['Expected'],
            color='blue', linewidth=2, label="Benford's Law")

    ax.set_title(f"Benford's Law: {input_file}", fontsize=16, weight='bold')
    ax.set_xlabel('First Digit', fontsize=14)
    ax.set_ylabel('Proportion', fontsize=14)
    ax.set_xticks(range(1, 10))
    ax.legend()
    plt.grid(True, which='major', linestyle='--', alpha=0.4)
    plt.tight_layout()
    plt.show()



# 🧮 Convert digit list to DataFrame
def digit_df(digits, label):
    df = pd.Series(digits).value_counts().sort_index().reset_index()
    df.columns = ['Digit', 'Freq']
    df['Prop'] = df['Freq'] / df['Freq'].sum()
    df['Type'] = label
    return pd.merge(df, benford_dist, on='Digit', how='left')




# 🔍 First-digit extraction
def get_first_digit(arr):
    return [int(str(int(np.floor(x)))[0]) for x in arr if x > 0]



def msg_blurb_square(msg, color):
    border = f"+{'-' * (len(msg) + 2)}+"
    print(f"{color}{border}\n| {msg} |\n{border}{color_reset}")

def read_xlsx(file_path):
    """Read data from an XLSX file and return a list of dictionaries."""
    wb = load_workbook(file_path)
    ws = wb.active
    headers = [cell.value for cell in ws[1]]
    data = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        row_data = dict(zip(headers, row))
        row_data["zero"] = "bobs your uncle" if row_data.get("zero") == "test" else row_data.get("zero")
        row_data["one"] = "avocadoes rule" if row_data.get("one") == "TeslaSucks" else row_data.get("one")
        data.append(row_data)
    
    wb.close()
    return data

    
def usage():
    print(f"Usage: {sys.argv[0]} -b [-I input.xlsx]")
    print("Example:")
    print(f"    {sys.argv[0]} -b")
    print(f"    {sys.argv[0]} -b -I benfordsLaw_tester.xlsx")

if __name__ == '__main__':
    main()


# <<<<<<<<<<<<<<<<<<<<<<<<<< Revision History >>>>>>>>>>>>>>>>>>>>>>>>>>

"""


"""


# <<<<<<<<<<<<<<<<<<<<<<<<<< Future Wishlist  >>>>>>>>>>>>>>>>>>>>>>>>>>

"""



"""


# <<<<<<<<<<<<<<<<<<<<<<<<<<      notes            >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
The Manipulated Data section simulates financial values that are artificially rounded 
to the nearest thousand, creating unnaturally uniform distributions. These values do not 
follow the logarithmic pattern predicted by Benford’s Law, which typically governs organic 
datasets. By comparing the first-digit frequency of this manipulated data to Benford's 
expected distribution, the model helps demonstrate how fabricated or tampered numbers 
diverge from statistical norms.


The Authentic Data section contains numerical values that are naturally occurring 
and statistically organic. These figures are either real-world samples or generated 
using randomization techniques that closely mimic legitimate datasets. They tend to 
follow Benford’s Law, which predicts the frequency of first digits in many 
naturally formed datasets.


git clone https://github.com/SheetJS/enron_xls.git

"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      The End        >>>>>>>>>>>>>>>>>>>>>>>>>>
    