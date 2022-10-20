# Convert Autosys JIL <-> Excel

Convert Autosys' JIL file to Excel file OR Excel to JIL
I am not a developer.
It's very sloppy because I made it myself

## Requirement

- python
- Sloppy (It's very sloppy because I made it myself)
  -> Several additional packages are used inside

## Run the Script

1. JIL : autorep -J ALL -q > ALLJOB.jil
2. Excute the Following
   -> python convert_jil.py -J ALLJOB.jil (Converted ALLJOB.xlsx file is generated when excute)
3. copy ALLJOB.xlsx C_ALLJOB.xlsx
4. Excute the Following
   -> python convert_excel.py -E C_ALLJOB.xlsx (Converted C_ALLJOB.jil file is generated when excute)
