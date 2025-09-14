#!/usr/bin/env python3
"""
Dashboard Data Generator
Uses the existing ms_expenses_analysis.py functions to create JSON data for web dashboard.
"""

import json
import sys
from datetime import datetime
import traceback

# Import your existing functions
try:
    from ms_expenses_analysis import extract_month_totals, extract_category_data
except ImportError as e:
    print(f"Error importing analysis functions: {e}")
    sys.exit(1)

def create_dashboard_data():
    """Generate JSON data for the web dashboard using existing analysis functions."""
    try:
        print("Generating dashboard data...")
        
        # Use your existing functions to extract data
        monthly_df = extract_month_totals('MS Expenses.xlsx')
        expenses_df = extract_category_data('MS Expenses.xlsx')
        
        if monthly_df.empty:
            print("Warning: No monthly data found!")
            return
        
        # Create dashboard data structure
        dashboard_data = {
            'monthly_spending': {
                'labels': monthly_df['Sheet'].tolist(),
                'values': [float(x) for x in monthly_df['Total'].tolist()],  # Ensure JSON serializable
                'months': monthly_df['Month'].tolist()
            },
            'last_updated': datetime.now().isoformat(),
            'summary': {
                'total': float(monthly_df['Total'].sum()),
                'average': float(monthly_df['Total'].mean()),
                'median': float(monthly_df['Total'].median()),
                'count': len(monthly_df),
                'highest_month': monthly_df.loc[monthly_df['Total'].idxmax(), 'Sheet'],
                'highest_amount': float(monthly_df['Total'].max()),
                'lowest_month': monthly_df.loc[monthly_df['Total'].idxmin(), 'Sheet'],
                'lowest_amount': float(monthly_df['Total'].min()),
            }
        }
        
        # Add category data if available
        if not expenses_df.empty:
            category_totals = expenses_df.groupby('Category')['Amount'].sum().sort_values(ascending=False)
            dashboard_data['category_spending'] = {
                'labels': category_totals.index.tolist(),
                'values': [float(x) for x in category_totals.values.tolist()]
            }
            
            # Add monthly category breakdown
            monthly_categories = expenses_df.groupby(['Month', 'Category'])['Amount'].sum().unstack(fill_value=0)
            if not monthly_categories.empty:
                dashboard_data['monthly_categories'] = {
                    'months': monthly_categories.index.tolist(),
                    'categories': monthly_categories.columns.tolist(),
                    'data': monthly_categories.values.tolist()
                }
        
        # Save dashboard data as JSON
        with open('dashboard_data.json', 'w') as f:
            json.dump(dashboard_data, f, indent=2, default=str)
        
        print(f"✅ Dashboard data generated successfully!")
        print(f"   - {len(monthly_df)} months of data")
        print(f"   - Total spending: ${dashboard_data['summary']['total']:,.2f}")
        print(f"   - Average monthly: ${dashboard_data['summary']['average']:,.2f}")
        
        if not expenses_df.empty:
            print(f"   - {len(expenses_df)} individual expenses")
            print(f"   - {len(category_totals)} categories")
        
        return dashboard_data
        
    except Exception as e:
        print(f"❌ Error generating dashboard data: {e}")
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    create_dashboard_data()