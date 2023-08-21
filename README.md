# Set-up and opening the file

### Step 1
Download, extract and paste the Python files in **C:\TOS_SF_Data_Processing**

### Step 2
Open Visual Studio Code, go to **File > Open Folder (Ctrl+K Ctrl+O)** and select the **TOS_OAS_Enhancement folder**

### Step 3
Go to Terminal > New Terminal and enter
```python
pip install -r requirements.txt
```

# How to use

### Step 1
Run the first python file and select the Excel file you want to process

### Step 2
Manual insert the relevant pivot tables in the Check sheet

### Step 3
Manual edit the CSV sheets in the Excel file, upload the CSVs with Data Loader and paste all relevant info into the Mappings tab

### Step 4
Save and close the Excel file, then run the second python file and select the same Excel file

### Step 5
In the 4_Working sheet, change the Course Paper and insert the Seq Number accordingly

### Step 6
Manual edit the last 4_CSV sheet, upload the 4_CSV sheet with Data Loader and you're done
