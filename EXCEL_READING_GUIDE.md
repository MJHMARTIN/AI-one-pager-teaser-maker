# Excel Reading Guide: Row-by-Row Processing

## ✅ Confirmation: Excel IS Read Row-by-Row (Left to Right)

Your Excel files **ARE being read correctly** from left to right, row by row. Here's the proof:

### How Row-Based Excel Works

For Excel files with "Label" and "Value" columns:

```
| Label                                    | Value                              |
|------------------------------------------|------------------------------------|
| Sponsor Name                             | GreenEnergy Capital Partners       |
| Primary Focus Area                       | Renewable Energy & Infrastructure  |
| Country of Incorporation                 | United States                      |
```

The system reads **LEFT TO RIGHT** for each row:
- Row 1: "Sponsor Name" (left) → "GreenEnergy Capital Partners" (right)
- Row 2: "Primary Focus Area" (left) → "Renewable Energy & Infrastructure" (right)
- Row 3: "Country of Incorporation" (left) → "United States" (right)

This creates a dictionary:
```python
{
  'sponsor name': 'GreenEnergy Capital Partners',
  'primary focus area': 'Renewable Energy & Infrastructure',
  'country of incorporation': 'United States',
  ...
}
```

## How to Diagnose Your Excel File

Run the diagnostic tool to verify your Excel is being read correctly:

```bash
python diagnose_excel.py your_file.xlsx
```

This will show you:
1. ✅ How many fields were parsed
2. ✅ What module type was detected
3. ✅ Which :TAG: placeholders will work
4. ✅ What values are available

## Module-Specific Tags

The system automatically detects which module type your Excel represents:

### Module1 (Sponsor-based)
Use these tags in your PowerPoint:
- `:ISSUER:` → maps to "Sponsor Name"
- `:INDUSTRY:` → maps to "Primary Focus Area"
- `:JURISDICTION:` → maps to "Country of Incorporation"
- `:ISSUANCE_TYPE:` → maps to "Financing Type"
- `:INITIAL_NOTIONAL:` → maps to "Total Financing Amount"
- `:COUPON_RATE:` → maps to "Coupon Rate"
- `:COUPON_FREQUENCY:` → maps to "Coupon Frequency"
- `:TENOR:` → maps to "Requested Tenor"
- `:CLIENT_SUMMARY:` → maps to "Sponsor Summary"
- `:PROJECT_HIGHLIGHT:` → maps to "Sponsor Background Investment Strategy"

### Module2 (Company-based)
Use these tags in your PowerPoint:
- `:ISSUER:` → maps to "Company Legal Name"
- `:INDUSTRY:` → maps to "Primary Industry"
- `:JURISDICTION:` → maps to "Country of Incorporation"
- `:ISSUANCE_TYPE:` → maps to "Financing Type"
- `:INITIAL_NOTIONAL:` → maps to "Total Financing Amount"
- `:COUPON_RATE:` → maps to "Coupon Rate"
- `:COUPON_FREQUENCY:` → maps to "Coupon Frequency"
- `:TENOR:` → maps to "Requested Tenor"
- `:CLIENT_SUMMARY:` → maps to "Company Overview Business Model"
- `:PROJECT_HIGHLIGHT:` → maps to "Company Growth Strategy Financial Projections"

### Module3 (Project-based)
Use these tags in your PowerPoint:
- `:ISSUER:` → maps to "Project Name"
- `:INDUSTRY:` → maps to "Project Type"
- `:JURISDICTION:` → maps to "Project Location Country"
- `:ISSUANCE_TYPE:` → maps to "Financing Type"
- `:INITIAL_NOTIONAL:` → maps to "Total Project Cost"
- `:COUPON_RATE:` → maps to "Coupon Rate"
- `:COUPON_FREQUENCY:` → maps to "Coupon Frequency"
- `:TENOR:` → maps to "Project Tenor"
- `:CLIENT_SUMMARY:` → maps to "Project Description"
- `:PROJECT_HIGHLIGHT:` → maps to "Project Overview Technical Specs Impact"

## Common Issues & Solutions

### Issue: "Placeholders are coming up empty"

**Solution 1: Check your placeholder format**
- ✅ Correct for row-based Excel: `:ISSUER:`, `:INDUSTRY:`, `:TENOR:`
- ❌ Incorrect: `[ISSUER]`, `{ISSUER}`, `ISSUER`

**Solution 2: Verify module type**
Run the diagnostic tool to see which module was detected. Make sure you're using the correct tags for that module type.

**Solution 3: Check Excel structure**
Your Excel should have exactly 2 columns:
- Column A: Labels (e.g., "Sponsor Name", "Industry", etc.)
- Column B: Values

**Solution 4: Use flexible label names**
The system normalizes labels, so these all work the same:
- "Sponsor Name" = "sponsor name" = "SPONSOR NAME" = "Sponsor  Name"
- Label variations are supported (e.g., "Company name" = "Name of company")

### Issue: "Some fields work but others don't"

**Check 1: Verify the field exists in your Excel**
```bash
python diagnose_excel.py your_file.xlsx
```
Look at the "ALL PARSED FIELDS" section to see what's available.

**Check 2: Verify you're using the right tag for your module**
Each module has different field names. :ISSUER: maps to:
- Module1: "Sponsor Name"
- Module2: "Company Legal Name"
- Module3: "Project Name"

### Issue: "Multi-sheet Excel not working"

For multi-sheet Excel files:
1. ✅ All sheets are read automatically
2. ✅ Each sheet's row-based data is parsed left-to-right
3. ✅ All data is merged into one unified dictionary
4. Optional: Enable "sheet namespacing" to prefix fields with sheet names

## Hybrid Mode (Automatic)

The system uses "hybrid mode" which reads **BOTH**:
1. **Column-based**: First row as headers, second row as values
2. **Row-based**: First column as labels, second column as values

This means:
- ✅ If your Excel has "Label/Value" columns, it's read row-by-row
- ✅ If your Excel has column headers, they're also read
- ✅ Both formats work simultaneously

### Important Fix Applied

**Previous issue**: The "Label" and "Value" column headers were being added as data fields, causing confusion.

**Fixed**: Now detects when columns are named "Label"/"Value" and skips adding them as data fields. Only the actual row-by-row data is used.

## Troubleshooting Commands

### 1. Diagnose your Excel file
```bash
python diagnose_excel.py your_file.xlsx
```

### 2. Test in the Streamlit app
When you upload your Excel in the web interface:
- Look for "Hybrid mode: Using both column headers and row-based data"
- Click "🔍 View parsed fields" to see all extracted data
- Check the module type detection message

### 3. Check application logs
When running the app, watch for these log messages:
- "Detected module type: Module1" (or Module2, Module3)
- "Parsed X fields from Excel"
- "Resolved :TAG: ..." messages show which tags are working

## Summary

✅ **Excel IS being read row-by-row (left to right)**
- First column = Labels
- Second column = Values

✅ **The system automatically:**
- Detects your module type (Module1, Module2, or Module3)
- Normalizes field names (case-insensitive, punctuation-insensitive)
- Maps :TAG: placeholders to the correct Excel fields
- Supports fuzzy matching for label variations

✅ **If placeholders are empty:**
1. Run the diagnostic tool: `python diagnose_excel.py your_file.xlsx`
2. Verify you're using :TAG: format (not [TAG])
3. Check that tags match your detected module type
4. Verify the field exists in your Excel

## Example Working Configuration

**Excel file (test_module1_sponsor.xlsx):**
```
Label                                    | Value
---------------------------------------- | ----------------------------------
Sponsor Name                             | GreenEnergy Capital Partners
Primary Focus Area                       | Renewable Energy & Infrastructure
Country of Incorporation                 | United States
```

**PowerPoint template:**
```
:ISSUER:        → Fills with "GreenEnergy Capital Partners"
:INDUSTRY:      → Fills with "Renewable Energy & Infrastructure"
:JURISDICTION:  → Fills with "United States"
```

**Diagnostic output:**
```
✅ Parsed 10 fields
✅ Detected: Module1
✅ :ISSUER: → 'GreenEnergy Capital Partners'
✅ :INDUSTRY: → 'Renewable Energy & Infrastructure'
✅ :JURISDICTION: → 'United States'
```

This configuration is **working perfectly** with row-by-row reading from left to right!
