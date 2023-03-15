using ElectricalEngineering
using Unitful, Unitful.DefaultSymbols
using WorkflowXLSX
using Format
using Dates

# Create string with actual date and time
datetime = Dates.format(now(), "yyyy-mm-dd_HHMMSS")

# Input XLSX file
infile = "Ohm.xlsx"
# Output XLSX file
outfile = "Ohm_result.xlsx"
# Make a copy of from the input to the output file
cp(infile, outfile, force=true)

# Rename first sheet with actual date and time
XLSXsheetname(outfile, datetime)

# Read voltage and resistance
V1 = XLSXreadvar(infile, "V1")
R1 = XLSXreadvar(infile, "R1")
# Calculate current
I1 = V1 / R1
# Write current to output file
XLSXwritevar(outfile, "I1", I1, A)
