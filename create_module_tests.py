import pandas as pd

# Module 1: Sponsor-focused
module1_data = {
    "Label": [
        "Sponsor Name",
        "Primary Focus Area",
        "Country of Incorporation",
        "Financing Type",
        "Total Financing Amount",
        "Coupon Rate",
        "Coupon Frequency",
        "Requested Tenor",
        "Sponsor Summary",
        "Sponsor Background Investment Strategy"
    ],
    "Value": [
        "GreenEnergy Capital Partners",
        "Renewable Energy & Infrastructure",
        "United States",
        "Senior Secured Notes",
        "$500M",
        "6.75%",
        "Semi-Annual",
        "7 years",
        "Leading renewable energy investment firm with $5B+ AUM",
        "GreenEnergy Capital Partners specializes in large-scale renewable projects across North America. Their investment strategy focuses on proven technologies with long-term power purchase agreements."
    ]
}

# Module 2: Company-focused
module2_data = {
    "Label": [
        "Company Legal Name",
        "Primary Industry",
        "Country of Incorporation",
        "Financing Type",
        "Total Financing Amount",
        "Coupon Rate",
        "Coupon Frequency",
        "Requested Tenor",
        "Company Overview Business Model",
        "Company Growth Strategy Financial Projections"
    ],
    "Value": [
        "SolarTech Industries Inc.",
        "Solar Panel Manufacturing",
        "Germany",
        "Convertible Bonds",
        "€300M",
        "4.5%",
        "Annual",
        "5 years",
        "Leading European solar panel manufacturer with 15% market share and vertically integrated production facilities",
        "SolarTech plans to expand manufacturing capacity by 40% over next 3 years. Revenue projected to grow from €800M to €1.2B with EBITDA margins improving from 18% to 22%."
    ]
}

# Module 3: Project-focused
module3_data = {
    "Label": [
        "Project Name",
        "Project Type",
        "Project Location Country",
        "Financing Type",
        "Total Project Cost",
        "Coupon Rate",
        "Coupon Frequency",
        "Project Tenor",
        "Project Description",
        "Project Overview Technical Specs Impact"
    ],
    "Value": [
        "Sunrise Wind Farm Phase II",
        "Offshore Wind",
        "United Kingdom",
        "Project Finance",
        "£450M",
        "5.25%",
        "Quarterly",
        "10 years",
        "500MW offshore wind facility located 25km off the coast of Scotland",
        "The project features 75 next-generation 7MW turbines with advanced monitoring systems. Expected to generate 1,800 GWh annually, powering 450,000 homes and reducing CO2 emissions by 900,000 tons per year."
    ]
}

# Create Excel files
df1 = pd.DataFrame(module1_data)
df2 = pd.DataFrame(module2_data)
df3 = pd.DataFrame(module3_data)

df1.to_excel("test_module1_sponsor.xlsx", index=False)
df2.to_excel("test_module2_company.xlsx", index=False)
df3.to_excel("test_module3_project.xlsx", index=False)

print("✅ Created test files:")
print("  - test_module1_sponsor.xlsx")
print("  - test_module2_company.xlsx")
print("  - test_module3_project.xlsx")
