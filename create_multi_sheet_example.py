#!/usr/bin/env python3
"""Create a multi-sheet Excel file for testing."""

import pandas as pd

# Sheet 1: Company Information
company_data = {
    'Label': [
        'Company name',
        'Sector',
        'Company Operations',
        'Industry',
    ],
    'Value': [
        'GreenTech Solutions',
        'New Energy',
        'develops and operates renewable energy solutions, specializing in utility-scale installations and promotes clean energy transition initiatives',
        'Renewable Energy',
    ]
}

# Sheet 2: Project Details
project_data = {
    'Label': [
        'Project Focus',
        'Project Type',
        'Country',
        'Location State',
        'Technology',
        'Unit',
        'Project Status',
        'COD',
    ],
    'Value': [
        'development of a 300 MW utility-scale wind farm with integrated battery storage in Australia',
        'Hybrid renewable energy facility',
        'Australia',
        'Queensland',
        'Wind turbines with 150 MWh battery storage',
        '300 MW',
        'Development Stage',
        'Q2 2028',
    ]
}

# Sheet 3: Partnerships & Scope
partnership_data = {
    'Label': [
        'Service Scope',
        'Partnership',
        'Partnership Benefit',
        'EPC Contractor',
        'Offtaker',
    ],
    'Value': [
        'provides end-to-end support from site assessment and land acquisition to engineering, procurement, construction, and integrated operations',
        'partnered with leading turbine manufacturers and energy storage providers in the Asia-Pacific region',
        'This collaboration establishes a resilient hybrid energy ecosystem ensuring grid stability and optimized renewable dispatch',
        'WindTech International',
        'Queensland Energy Grid',
    ]
}

# Sheet 4: Financial Information
financial_data = {
    'Label': [
        'Initial Investment',
        'Expansion Investment',
        'Expansion Path',
        'Total Project Cost',
    ],
    'Value': [
        '$250-300 M USD',
        '$1.5 B USD',
        'across 6 additional renewable projects in Australia and New Zealand',
        '$1.8 B USD',
    ]
}

# Create Excel writer
with pd.ExcelWriter('multi_sheet_example.xlsx', engine='openpyxl') as writer:
    pd.DataFrame(company_data).to_excel(writer, sheet_name='Company Info', index=False)
    pd.DataFrame(project_data).to_excel(writer, sheet_name='Project Details', index=False)
    pd.DataFrame(partnership_data).to_excel(writer, sheet_name='Partnerships', index=False)
    pd.DataFrame(financial_data).to_excel(writer, sheet_name='Financials', index=False)

print("✅ Created multi_sheet_example.xlsx")
print("\n📚 Sheets:")
print("  1. Company Info - Company details")
print("  2. Project Details - Project specifications")
print("  3. Partnerships - Partnerships and scope")
print("  4. Financials - Investment information")
print("\n💡 All data will be merged into one unified knowledge base!")
