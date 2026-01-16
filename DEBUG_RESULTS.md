# Debug Results: Mapping Layer Investigation

## Test Results

✅ **Excel parsing: WORKING PERFECTLY**
- Test file `test_module1_sponsor.xlsx` parses to 10 fields
- All field names are correct: "sponsor name", "primary focus area", etc.
- No "unnamed X" fields

✅ **Module detection: WORKING PERFECTLY**
- Correctly detects Module1 (10/10 fields matched)

✅ **Field mapping: WORKING PERFECTLY**
- All tags resolve correctly:
  - `:ISSUER:` → "GreenEnergy Capital Partners"
  - `:INDUSTRY:` → "Renewable Energy & Infrastructure"
  - `:JURISDICTION:` → "United States"
  - etc.

## The Problem

Since the mapping works perfectly with test data, the issue must be one of these:

### 1. **Your actual Excel file structure is broken** (Most Likely)
From the logs, your file produces 16,377 "unnamed X" fields instead of proper field names.

**Proof**: The log shows:
```
data_dict has 16377 keys. First 10: ['unnamed 0', 'unnamed 1', 'unnamed 2', ...]
Module detection: score=0 (nothing matched)
```

### 2. **PPT template doesn't have the expected placeholders** (Possible)
- Template might have `{ISSUER}` or `[ISSUER]` instead of `:ISSUER:`
- Placeholders might be split across multiple text runs (formatting issue)
- Placeholders might have smart quotes or hidden characters

### 3. **Different file being used in actual run** (Possible)
- Testing with `test_module1_sponsor.xlsx` works
- But you're uploading a different file that has structural problems

## Next Steps

### Step 1: Check YOUR actual Excel file
Upload your Excel in the Streamlit app and look at:
- The "View parsed fields" section
- Does it show 10-50 fields with real names? ✅ Good
- Does it show 16,000+ fields named "unnamed X"? ❌ Your Excel is malformed

### Step 2: Verify your PPT template has correct placeholders
Open your template in PowerPoint and check:
- Select a text box
- Verify it says exactly `:ISSUER:` (with colons, no spaces)
- Not `{ISSUER}`, not `[ISSUER]`, not `: ISSUER:` (with space)

### Step 3: Try with test data
1. Use `test_module1_sponsor.xlsx` (proven working)
2. Create a simple PPT with just: `:ISSUER:` and `:INDUSTRY:`
3. If this works, your Excel is the problem
4. If this doesn't work, your PPT template is the problem

## Enhanced Logging Added

I've added detailed logging that will show:

1. **In app.py (line ~270)**:
   ```
   CRITICAL DEBUG: data_dict being passed to fill_pptx:
     Keys count: X
     First 10 keys: [...]
   ```

2. **In ppt_filler.py fill_pptx (line ~410)**:
   ```
   RECEIVED IN fill_pptx:
     data_dict: X keys
     First 10 keys: [...]
   ```

3. **In process_placeholder (line ~120)**:
   ```
   🔍 PROCESSING TAG: ISSUER
      data_dict has X keys
      placeholder_mapping has entry for ISSUER: True/False
   ```

4. **In resolve_tag_to_value (line ~37)**:
   ```
   🔍 resolve_tag_to_value called for: ISSUER
      Module mapping: ISSUER → 'sponsor name'
      ✅ FOUND in data_dict: 'GreenEnergy Capital Partners'
   ```

## To See Detailed Logs

After you try to generate a PPT:
```bash
tail -100 streamlit.log
```

Look for:
- "CRITICAL DEBUG" sections showing what data was passed
- "🔍 PROCESSING TAG" showing each tag being resolved
- "✅ FOUND" or "❌ NOT FOUND" for each resolution
- If you see "❌ NOT FOUND", the next line shows what keys ARE available

## Expected Good Log Output

```
CRITICAL DEBUG: data_dict being passed to fill_pptx:
  Keys count: 10
  First 10 keys: ['sponsor name', 'primary focus area', ...]

🔍 PROCESSING TAG: ISSUER
   Module mapping: ISSUER → 'sponsor name' (normalized: 'sponsor name')
   ✅ FOUND in data_dict: 'GreenEnergy Capital Partners'
```

## Expected Bad Log Output (Your Current Problem)

```
CRITICAL DEBUG: data_dict being passed to fill_pptx:
  Keys count: 16377
  First 10 keys: ['unnamed 0', 'unnamed 1', 'unnamed 2', ...]

🔍 PROCESSING TAG: ISSUER
   Module mapping: ISSUER → 'sponsor name' (normalized: 'sponsor name')
   ❌ NOT FOUND: 'sponsor name' not in data_dict keys
   Available keys: ['unnamed 0', 'unnamed 1', ...]
```

## The Fix

**If it's your Excel**: Restructure it to match the test file format:
```
| Label                    | Value                        |
|--------------------------|------------------------------|
| Sponsor Name             | Your Company Name            |
| Primary Focus Area       | Your Industry                |
| Country of Incorporation | Your Country                 |
| Financing Type           | Your Type                    |
| Total Financing Amount   | $500M                        |
| Coupon Rate              | 6.75%                        |
| Coupon Frequency         | Semi-Annual                  |
| Requested Tenor          | 7 years                      |
| Sponsor Summary          | Your summary...              |
| Sponsor Background...    | Your background...           |
```

**If it's your PPT**: Make sure placeholders are:
- Exactly `:ISSUER:`, `:INDUSTRY:`, etc. (with colons)
- Not split across multiple text styles
- Not using smart quotes or special characters

## Verification

Run this on YOUR actual Excel file:
```bash
python3 test_mapping.py
# Replace the file name in the script with yours
```

If that test passes but the app still shows empty:
→ The problem is in your PPT template, not the data

If that test fails with "unnamed X" keys:
→ The problem is your Excel structure
