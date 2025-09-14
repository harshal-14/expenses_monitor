#!/usr/bin/env python3
"""
MS Expenses Analysis Script
Analyzes monthly spending data from MS Expenses.xlsx with multiple month sheets.
Creates visualizations of spending totals by month and category breakdown.
"""

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from datetime import datetime
import re
from pathlib import Path
import warnings
warnings.filterwarnings('ignore')

# Set up plotting style
plt.style.use('seaborn-v0_8')
sns.set_palette("husl")

def extract_month_totals(excel_file='MS Expenses.xlsx'):
    """Extract total spending from each month sheet."""
    xl = pd.ExcelFile(excel_file)
    monthly_data = []
    
    print(f"Found {len(xl.sheet_names)} sheets: {xl.sheet_names}")
    
    for sheet_name in xl.sheet_names:
        # Skip non-month sheets
        if sheet_name.lower() in ['loan repayment']:
            continue
            
        print(f"Processing sheet: {sheet_name}")
        
        # Read the sheet
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        
        # Look for "Total" anywhere in the dataframe
        total_value = None
        found_location = None
        
        # Search through all cells in the dataframe
        for row_idx in range(len(df)):
            for col_name in df.columns:
                cell_value = str(df.iloc[row_idx, df.columns.get_loc(col_name)])
                
                # Check if this cell contains "Total" (case insensitive)
                if 'total' in cell_value.lower() and cell_value.lower() != 'nan':
                    print(f"  Found 'Total' at row {row_idx}, column '{col_name}': {cell_value}")
                    
                    # Look for numeric values in the same row (all columns)
                    for check_col in df.columns:
                        try:
                            potential_value = df.iloc[row_idx, df.columns.get_loc(check_col)]
                            if pd.notna(potential_value) and isinstance(potential_value, (int, float)) and potential_value > 0:
                                # Skip values that are too small (likely not the total)
                                if potential_value > 10:  # Reasonable threshold for monthly totals
                                    if not total_value or potential_value > total_value:  # Take the largest reasonable value
                                        total_value = potential_value
                                        found_location = f"row {row_idx}, col '{check_col}'"
                                        print(f"    Found potential total: ${potential_value:.2f} at {found_location}")
                        except (ValueError, TypeError):
                            continue
                    
                    # Also check adjacent cells (next column, next row)
                    # Check next column
                    col_idx = df.columns.get_loc(col_name)
                    if col_idx + 1 < len(df.columns):
                        next_col = df.columns[col_idx + 1]
                        try:
                            potential_value = df.iloc[row_idx, col_idx + 1]
                            if pd.notna(potential_value) and isinstance(potential_value, (int, float)) and potential_value > 10:
                                if not total_value or potential_value > total_value:
                                    total_value = potential_value
                                    found_location = f"row {row_idx}, col '{next_col}'"
                                    print(f"    Found potential total (next col): ${potential_value:.2f} at {found_location}")
                        except (ValueError, TypeError, IndexError):
                            pass
                    
                    # Check next row, same column
                    if row_idx + 1 < len(df):
                        try:
                            potential_value = df.iloc[row_idx + 1, df.columns.get_loc(col_name)]
                            if pd.notna(potential_value) and isinstance(potential_value, (int, float)) and potential_value > 10:
                                if not total_value or potential_value > total_value:
                                    total_value = potential_value
                                    found_location = f"row {row_idx + 1}, col '{col_name}'"
                                    print(f"    Found potential total (next row): ${potential_value:.2f} at {found_location}")
                        except (ValueError, TypeError, IndexError):
                            pass
        
        # If still no total found, try to find the largest reasonable numeric value
        if not total_value:
            print(f"  No 'Total' text found, looking for largest numeric value...")
            numeric_values = []
            for row_idx in range(len(df)):
                for col_name in df.columns:
                    try:
                        val = df.iloc[row_idx, df.columns.get_loc(col_name)]
                        if pd.notna(val) and isinstance(val, (int, float)) and val > 10:
                            numeric_values.append(val)
                    except:
                        continue
            
            if numeric_values:
                # Take the largest value that seems reasonable (not absurdly large)
                numeric_values.sort(reverse=True)
                for val in numeric_values:
                    if val < 10000:  # Reasonable upper bound for monthly spending
                        total_value = val
                        found_location = "largest numeric value"
                        break
        
        if total_value:
            monthly_data.append({
                'Sheet': sheet_name,
                'Month': normalize_month_name(sheet_name),
                'Total': total_value
            })
            print(f"  ✓ Selected total: ${total_value:.2f} (from {found_location})")
        else:
            print(f"  ✗ No total found for {sheet_name}")
    
    return pd.DataFrame(monthly_data)

def normalize_month_name(sheet_name):
    """Normalize month names for consistent sorting."""
    # Mapping for month normalization
    month_mapping = {
        'may': '2024-05',
        'june': '2024-06', 
        'july': '2024-07',
        'august': '2024-08',
        'september': '2024-09',
        'october': '2024-10',
        'oct-november': '2024-11',
        'december-jan': '2024-12',
        'january-feb': '2025-01',
        'feb-march': '2025-02',
        'april25': '2025-04',
        'may25': '2025-05',
        'june25': '2025-06',
        'july25': '2025-07',
        'aug25': '2025-08',
        'sep25': '2025-09'
    }
    
    sheet_lower = sheet_name.lower()
    return month_mapping.get(sheet_lower, sheet_name)

def extract_category_data(excel_file='MS Expenses.xlsx'):
    """Extract spending data by category from all sheets."""
    xl = pd.ExcelFile(excel_file)
    all_expenses = []
    
    # Define category mappings based on common expense types
    category_mapping = {
        'rent': 'Housing',
        'utilities': 'Housing', 
        'electricity': 'Utilities',
        'walmart': 'Shopping',
        'amazon': 'Shopping',
        'price chopper': 'Groceries',
        'market basket': 'Groceries',
        'apna bazaar': 'Groceries',
        'indian basket': 'Groceries',
        'uber': 'Transportation',
        'commuter rail': 'Transportation',
        'scooter': 'Transportation',
        'dragon dynasty': 'Restaurants',
        'sebastian office meal': 'Restaurants',
        'uber eats': 'Restaurants',
        'pizza': 'Restaurants',
        'mela': 'Entertainment',
        'whale watch': 'Entertainment',
        'zolve': 'Banking/Finance',
        'prime': 'Subscriptions'
    }
    
    for sheet_name in xl.sheet_names:
        if sheet_name.lower() in ['loan repayment']:
            continue
            
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        month = normalize_month_name(sheet_name)
        
        # Extract individual expenses
        if 'Item' in df.columns and 'Cost' in df.columns:
            for _, row in df.iterrows():
                item = str(row['Item']).lower() if pd.notna(row['Item']) else ''
                cost = row['Cost'] if pd.notna(row['Cost']) else 0
                
                if item and cost > 0 and 'total' not in item:
                    # Categorize the expense
                    category = 'Other'
                    for keyword, cat in category_mapping.items():
                        if keyword in item:
                            category = cat
                            break
                    
                    all_expenses.append({
                        'Month': month,
                        'Item': row['Item'],
                        'Category': category,
                        'Amount': cost
                    })
    
    return pd.DataFrame(all_expenses)

def plot_monthly_spending(monthly_df, save_path='monthly_spending.png'):
    """Create a plot of total spending per month."""
    # Sort by month
    monthly_df = monthly_df.sort_values('Month')
    
    plt.figure(figsize=(14, 8))
    bars = plt.bar(range(len(monthly_df)), monthly_df['Total'], 
                   color='steelblue', alpha=0.7, edgecolor='navy', linewidth=1.5)
    
    plt.title('Monthly Spending Totals', fontsize=18, fontweight='bold', pad=20)
    plt.xlabel('Month', fontsize=14)
    plt.ylabel('Amount ($)', fontsize=14)
    
    # Set x-axis labels
    plt.xticks(range(len(monthly_df)), monthly_df['Sheet'], rotation=45, ha='right')
    
    # Add value labels on bars
    for i, (bar, amount) in enumerate(zip(bars, monthly_df['Total'])):
        height = bar.get_height()
        plt.text(bar.get_x() + bar.get_width()/2., height + height*0.01,
                f'${amount:,.0f}', ha='center', va='bottom', fontsize=10, fontweight='bold')
    
    # Add average line
    avg_spending = monthly_df['Total'].mean()
    plt.axhline(y=avg_spending, color='red', linestyle='--', alpha=0.7, linewidth=2)
    plt.text(len(monthly_df)-1, avg_spending + avg_spending*0.05, 
             f'Average: ${avg_spending:,.0f}', ha='right', va='bottom', 
             color='red', fontweight='bold')
    
    plt.grid(axis='y', alpha=0.3)
    plt.tight_layout()
    plt.savefig(save_path, dpi=300, bbox_inches='tight')
    plt.show()
    
    return monthly_df

def plot_category_spending(expenses_df, save_path='category_spending.png'):
    """Create a plot of spending by category."""
    if expenses_df.empty:
        print("No expense data available for category analysis")
        return
        
    category_totals = expenses_df.groupby('Category')['Amount'].sum().sort_values(ascending=False)
    
    plt.figure(figsize=(12, 8))
    colors = sns.color_palette("husl", len(category_totals))
    bars = plt.barh(category_totals.index, category_totals.values, color=colors)
    
    plt.title('Total Spending by Category', fontsize=16, fontweight='bold')
    plt.xlabel('Amount ($)', fontsize=12)
    plt.ylabel('Category', fontsize=12)
    
    # Add value labels on bars
    for bar, amount in zip(bars, category_totals.values):
        width = bar.get_width()
        plt.text(width + width*0.01, bar.get_y() + bar.get_height()/2.,
                f'${amount:,.0f}', ha='left', va='center', fontsize=10)
    
    plt.grid(axis='x', alpha=0.3)
    plt.tight_layout()
    plt.savefig(save_path, dpi=300, bbox_inches='tight')
    plt.show()
    
    return category_totals

def plot_spending_trend(monthly_df, save_path='spending_trend.png'):
    """Create a line plot showing spending trend over time."""
    monthly_df = monthly_df.sort_values('Month')
    
    plt.figure(figsize=(14, 8))
    plt.plot(range(len(monthly_df)), monthly_df['Total'], 
             marker='o', linewidth=3, markersize=8, color='steelblue')
    
    plt.title('Spending Trend Over Time', fontsize=18, fontweight='bold', pad=20)
    plt.xlabel('Month', fontsize=14)
    plt.ylabel('Amount ($)', fontsize=14)
    
    plt.xticks(range(len(monthly_df)), monthly_df['Sheet'], rotation=45, ha='right')
    
    # Add trend line
    x_vals = np.arange(len(monthly_df))
    z = np.polyfit(x_vals, monthly_df['Total'], 1)
    p = np.poly1d(z)
    plt.plot(x_vals, p(x_vals), "--", alpha=0.8, color='red', linewidth=2)
    
    # Add annotations for min and max
    min_idx = monthly_df['Total'].idxmin()
    max_idx = monthly_df['Total'].idxmax()
    
    min_val = monthly_df.loc[min_idx, 'Total']
    max_val = monthly_df.loc[max_idx, 'Total']
    min_month = monthly_df.loc[min_idx, 'Sheet']
    max_month = monthly_df.loc[max_idx, 'Sheet']
    
    plt.annotate(f'Lowest: ${min_val:,.0f}', 
                xy=(list(monthly_df.index).index(min_idx), min_val),
                xytext=(10, 10), textcoords='offset points',
                bbox=dict(boxstyle='round,pad=0.3', fc='yellow', alpha=0.7),
                arrowprops=dict(arrowstyle='->', connectionstyle='arc3,rad=0'))
    
    plt.annotate(f'Highest: ${max_val:,.0f}', 
                xy=(list(monthly_df.index).index(max_idx), max_val),
                xytext=(10, -20), textcoords='offset points',
                bbox=dict(boxstyle='round,pad=0.3', fc='orange', alpha=0.7),
                arrowprops=dict(arrowstyle='->', connectionstyle='arc3,rad=0'))
    
    plt.grid(alpha=0.3)
    plt.tight_layout()
    plt.savefig(save_path, dpi=300, bbox_inches='tight')
    plt.show()

def generate_summary_report(monthly_df, expenses_df):
    """Generate a comprehensive summary report."""
    total_spending = monthly_df['Total'].sum()
    avg_monthly = monthly_df['Total'].mean()
    median_monthly = monthly_df['Total'].median()
    std_monthly = monthly_df['Total'].std()
    
    print("\n" + "="*60)
    print("MS EXPENSES ANALYSIS SUMMARY")
    print("="*60)
    print(f"Analysis Period: {len(monthly_df)} months")
    print(f"Total Spending: ${total_spending:,.2f}")
    print(f"Average Monthly Spending: ${avg_monthly:,.2f}")
    print(f"Median Monthly Spending: ${median_monthly:,.2f}")
    print(f"Monthly Spending Std Dev: ${std_monthly:,.2f}")
    
    # Find highest and lowest months
    max_month = monthly_df.loc[monthly_df['Total'].idxmax()]
    min_month = monthly_df.loc[monthly_df['Total'].idxmin()]
    
    print(f"\nHighest Spending Month: {max_month['Sheet']} (${max_month['Total']:,.2f})")
    print(f"Lowest Spending Month: {min_month['Sheet']} (${min_month['Total']:,.2f})")
    print(f"Spending Range: ${max_month['Total'] - min_month['Total']:,.2f}")
    
    if not expenses_df.empty:
        print(f"\nCategory Breakdown:")
        category_totals = expenses_df.groupby('Category')['Amount'].sum().sort_values(ascending=False)
        for category, amount in category_totals.head().items():
            percentage = (amount / category_totals.sum()) * 100
            print(f"  {category}: ${amount:,.2f} ({percentage:.1f}%)")
    
    print(f"\nMonthly Breakdown:")
    for _, row in monthly_df.sort_values('Month').iterrows():
        print(f"  {row['Sheet']}: ${row['Total']:,.2f}")

def main():
    """Main function to run the MS Expenses analysis."""
    print("MS Expenses Analysis Tool")
    print("=" * 30)
    
    excel_file = 'MS Expenses.xlsx'
    
    # Extract monthly totals
    print(f"\nExtracting monthly totals from {excel_file}...")
    monthly_df = extract_month_totals(excel_file)
    
    if monthly_df.empty:
        print("No monthly data found!")
        return
    
    # Extract category data
    print(f"\nExtracting category data...")
    expenses_df = extract_category_data(excel_file)
    
    # Create output directory
    output_dir = Path('ms_expenses_plots')
    output_dir.mkdir(exist_ok=True)
    
    # Generate visualizations
    print(f"\nGenerating visualizations...")
    
    plot_monthly_spending(monthly_df, output_dir / 'monthly_spending.png')
    plot_spending_trend(monthly_df, output_dir / 'spending_trend.png')
    
    if not expenses_df.empty:
        plot_category_spending(expenses_df, output_dir / 'category_spending.png')
    
    # Generate summary report
    generate_summary_report(monthly_df, expenses_df)
    
    print(f"\nAll plots saved to: {output_dir}")
    print("Analysis complete!")

if __name__ == "__main__":
    main()
