#!/usr/bin/env python3
"""Create sample row-based Excel files for testing the new system."""

import pandas as pd

# Create a row-based Excel file
data = {
    'Label': [
        'Company name',
        'Sponsor Name',
        'Location - Country',
        'Location State',
        'Asset Type',
        'Project Type',
        'Unit',
        'Technology',
        'Initial Investment',
        'Future Investment',
        'EPC Contractor',
        'Offtaker',
        'Project Status',
        'Commercial Operation Date',
        'Description',
        'Title',
        'Industry',
    ],
    'Value': [
        'CLimAIte',
        'Green Energy Partners',
        'Indonesia',
        'West Java',
        'Renewable Energy',
        'Greenfield Development',
        '150 MW',
        'Solar PV with Battery Storage',
        '$180 million',
        '$50 million',
        'SunTech International',
        'PT Indonesia Power',
        'Development Stage',
        'Q4 2027',
        'Large-scale solar facility with advanced energy storage systems designed to provide reliable baseload power',
        'CLimAIte Solar Project',
        'Renewable Energy',
    ]
}

df = pd.DataFrame(data)

# Save to Excel
df.to_excel('sample_row_based.xlsx', index=False)
print("✅ Created sample_row_based.xlsx")

# Show preview
print("\n📋 Preview:")
print(df.to_string(index=False))
