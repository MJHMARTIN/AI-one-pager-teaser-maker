#!/usr/bin/env python3
"""Create test Excel files with the user's example data."""

import pandas as pd

# Example 1: Battery Storage / Japan
example1_data = {
    'Label': [
        'Sector',
        'Country',
        'Project_Type',
        'Company_Operations',
        'Project_Focus',
        'Service_Scope',
        'Partnership',
        'Partnership_Benefit',
        'Initial_Investment',
        'Expansion_Investment',
        'Expansion_Path',
    ],
    'Value': [
        'New Energy',
        'Japan',
        'Grid-connected battery storage station',
        'develops and operates distributed energy systems, as well as develops and promotes the green transformation projects',
        'development of the grid-connected battery storage station project in Japan',
        'provides comprehensive support for all processes, from land acquisition to the construction and operation of battery storage stations',
        'partner with well-recognized partners, responsible for power market operations and the battery storage system',
        'This collaboration ensures a robust framework for the construction and operation of grid-scale battery storage facilities',
        '$3-30 M USD',
        '$80 M USD',
        'through repeat projects',
    ]
}

df1 = pd.DataFrame(example1_data)
df1.to_excel('example1_battery_japan.xlsx', index=False)
print("✅ Created example1_battery_japan.xlsx")

# Example 2: Solar / Vietnam
example2_data = {
    'Label': [
        'Sector',
        'Country',
        'Project_Type',
        'Company_Operations',
        'Project_Focus',
        'Service_Scope',
        'Partnership',
        'Partnership_Benefit',
        'Initial_Investment',
        'Expansion_Investment',
        'Expansion_Path',
    ],
    'Value': [
        'New Energy',
        'Vietnam',
        'Utility-scale solar photovoltaic farm',
        'develops and operates renewable energy solutions, specializing in utility-scale solar installations and promotes clean energy transition initiatives',
        'development of a 500 MW utility-scale solar photovoltaic farm in Vietnam',
        'provides end-to-end support from site selection and land acquisition to engineering, procurement, construction, and long-term operations and maintenance',
        'partnered with leading EPC contractors and grid operators in Southeast Asia',
        'This collaboration establishes a strong foundation for reliable grid integration and sustainable power generation operations',
        '$150-200 M USD',
        '$1.2 B USD',
        'across 5 additional solar projects in ASEAN region',
    ]
}

df2 = pd.DataFrame(example2_data)
df2.to_excel('example2_solar_vietnam.xlsx', index=False)
print("✅ Created example2_solar_vietnam.xlsx")

# Example 3: Wind / India
example3_data = {
    'Label': [
        'Sector',
        'Country',
        'Project_Type',
        'Company_Operations',
        'Project_Focus',
        'Service_Scope',
        'Partnership',
        'Partnership_Benefit',
        'Initial_Investment',
        'Expansion_Investment',
        'Expansion_Path',
    ],
    'Value': [
        'New Energy',
        'India',
        'Onshore wind farm with battery storage integration',
        'develops and operates onshore wind energy systems, as well as advances integrated wind-storage solutions for grid stabilization and energy transition acceleration',
        'development of a 250 MW onshore wind farm with 100 MWh battery storage integration in India',
        'provides comprehensive support across site assessment, environmental permitting, engineering design, EPC management, and integrated operations for wind and storage systems',
        'partner with Tier-1 turbine manufacturers and leading energy storage integrators',
        'This collaboration creates a resilient hybrid energy system ensuring grid reliability and optimized renewable energy dispatch',
        '$80-120 M USD',
        '$600 M USD',
        'through 4 additional hybrid renewable projects across Indian states',
    ]
}

df3 = pd.DataFrame(example3_data)
df3.to_excel('example3_wind_india.xlsx', index=False)
print("✅ Created example3_wind_india.xlsx")

print("\n📋 All test files created successfully!")
