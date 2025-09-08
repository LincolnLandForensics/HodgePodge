#!/usr/bin/env python3
# coding: utf-8
"""
Read and write XLSX files using openpyxl without Pandas.
Author: LincolnLandForensics
Version: 0.0.2
"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Imports        >>>>>>>>>>>>>>>>>>>>>>>>>>

import os
import sys
import argparse
# import openpyxl
# from openpyxl import load_workbook, Workbook
# from openpyxl.styles import PatternFill

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from scipy.stats import chisquare   # pip install scipy

# colors
color_red = color_yellow = color_green = color_blue = color_purple = color_reset = ''
from colorama import Fore, Back, Style
print(Back.BLACK)
color_red, color_yellow, color_green = Fore.RED, Fore.YELLOW, Fore.GREEN
color_blue, color_purple, color_reset = Fore.BLUE, Fore.MAGENTA, Style.RESET_ALL


# <<<<<<<<<<<<<<<<<<<<<<<<<<      Pre-Sets       >>>>>>>>>>>>>>>>>>>>>>>>>>

author = 'LincolnLandForensics'
description = "This script models Benford’s Law by generating and comparing authentic versus manipulated financial data. It outputs frequency distributions to Excel for forensic analysis, helping identify statistical anomalies suggestive of fraud."
version = '0.2.2'


# <<<<<<<<<<<<<<<<<<<<<<<<<<      Menu           >>>>>>>>>>>>>>>>>>>>>>>>>>



def main():
    global row
    row = 0  # defines arguments
    # Row = 1  # defines arguments   # if you want to add headers 
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument('-I', '--input', help='', required=False, default='benfordsLaw_tester.xlsx')
    # parser.add_argument('-O', '--output', help='', required=False)
    parser.add_argument('-b', '--benfords', help='read xlsx', required=False, action='store_true')
    # parser.add_argument('-c', '--column', help='choose column', required=False, action='store_true')
    parser.add_argument('-c', '--column', help='', required=False, default='A')
    args = parser.parse_args()


    global input_file
    # input_file = args.input if args.input else "benfordsLaw_tester.xlsx"
    input_file = args.input

    global column_pick
    # column_pick = args.column if args.column else "A"
    column_pick = args.column
    print(f'Reading______ {input_file} column______ {column_pick}')

    
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
    wb = load_workbook(input_file, data_only=True)
    ws = wb.active
    column_data = [cell.value for cell in ws[column_pick] if isinstance(cell.value, (int, float))]

    if not column_data:
        msg_blurb_square("No numeric data found in the selected Excel column. Skipping analysis.", color_yellow)
        return

    # Extract first digits from Excel data
    excel_first_digit = get_first_digit(column_data)

    # Benford expected distribution
    benford_dist = pd.DataFrame({
        'Digit': range(1, 10),
        'Expected': np.log10(1 + 1 / np.arange(1, 10))
    })

    # Convert digit list to DataFrame
    df = pd.Series(excel_first_digit).value_counts().sort_index().reset_index()
    df.columns = ['Digit', 'Freq']
    df['Prop'] = df['Freq'] / df['Freq'].sum()
    df['Type'] = 'Excel Data'
    df = pd.merge(df, benford_dist, on='Digit', how='left')

    # Chi-Square Test
    observed_counts = df.set_index('Digit').reindex(range(1, 10), fill_value=0)['Freq']
    expected_counts = benford_dist['Expected'] * len(excel_first_digit)
    chi2_stat, p_value = chisquare(f_obs=observed_counts, f_exp=expected_counts)

    # MAD Test
    mad = np.mean(np.abs(df['Prop'] - df['Expected']))

    # Plotting
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.bar(df['Digit'], df['Prop'], width=0.5, alpha=0.7, label='Excel Data', color='orange')
    ax.plot(benford_dist['Digit'], benford_dist['Expected'],
            color='blue', linewidth=2, label="Benford's Law")

    ax.set_title(f"Benford's Law Analysis: {input_file}", fontsize=16, weight='bold')
    ax.set_xlabel('First Digit', fontsize=14)
    ax.set_ylabel('Proportion', fontsize=14)
    ax.set_xticks(range(1, 10))
    ax.legend()
    plt.grid(True, which='major', linestyle='--', alpha=0.4)

    # Add Chi-Square and MAD results below the chart
    stats_text = f"Chi-Square: {chi2_stat:.2f} (p={p_value:.4f})   |   MAD: {mad:.4f}"
    fig.text(0.5, 0.01, stats_text,
             ha='center', va='bottom',
             fontsize=12,
             bbox=dict(boxstyle='round,pad=0.4', facecolor='#f0f0f0', edgecolor='gray'))

    plt.tight_layout(rect=[0, 0.03, 1, 1])  # Leave space at bottom for stats
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
    print(f"    {sys.argv[0]} -b -I benfordsLaw_tester.xlsx -c d")

if __name__ == '__main__':
    main()


# <<<<<<<<<<<<<<<<<<<<<<<<<< Revision History >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
0.1.0 - working copy

"""


# <<<<<<<<<<<<<<<<<<<<<<<<<< Future Wishlist  >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
if there are no instances of a number, like 3, it blows an error
    raise ValueError(f'shapes {shape1} and {shape2} could not be '
ValueError: shapes (8,) and (9,) could not be broadcast together


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

The Chi-Square Test is a statistical method used to evaluate whether observed frequencies 
in a dataset differ significantly from expected frequencies under a certain hypothesis. 
It's commonly applied to categorical data—like survey responses, classifications, or 
groupings—to test for independence or goodness-of-fit.
The bigger the number, the less likely it is that the differences occurred by chance, 
and the more likely it is that something meaningful is happening in the data.

Chi-Square Test ➤ Stat: 1114.61


git clone https://github.com/SheetJS/enron_xls.git

"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      The End        >>>>>>>>>>>>>>>>>>>>>>>>>>
    