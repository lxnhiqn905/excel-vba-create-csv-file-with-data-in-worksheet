# ==============================
# excel-vba-create-csv-file-with-data-in-worksheet
# Create csv file with 2 ways. 
# ==============================
1. Read and write in 1 times
  - Record limit: 30K lines
2. Read and write line by line.
  - Not limit records.

## Source code
  - [Read and write in 1 times.vba](./Read%20and%20write%20in%201%20times.vba)
  - [Read and write line by line.vba](./Read%20and%20write%20line%20by%20line.vba)

## How to use
  1. Create Excel enable Macro
  2. Create New Module and copy source code into that. Sample as file: [Create CSV - TestFile.xlsm](./Create%20CSV%20-%20TestFile.xlsm)
  3. Run ...
  4. Output as file [Read and write in 1 times.csv](./Read%20and%20write%20in%201%20times.csv) and [Read and write line by line.csv](./Read%20and%20write%20line%20by%20line.csv)

## Note: Need import Microsoft ActiveX Data Objects 6.1 Library to use
- Tool -> References -> Find Microsoft ActiveX Data Objects 6.1 Library and select it 
