# AuditSummary
This is a Python script to translate security audit data from Excel spreadsheets into a PDF suitable for an executive summary.

Using an Excel spreadsheet to complete audits is efficient but can be a pain to get the data exfilled to a summary or aggregated/easy-to-read format. This program has been built to use Windows systems to pull down all Excel files in a directory and read the contents from specified columns in a row range and print all data to a CSV. Then it will read the CSV to create a summarized view of the CSV.

## Note
Please pay close attention to the rows listed in the code as Excel, pwrsh, or Python may not read the rows as intended. You may have to make adjustments after running the script a few times.

## You may have to create some directories manually and alter the path in the code ##

### Installation
# These steps are for installing but there is no way this will run... LOL
# I have placed it here simply for you to use to help build yours out

1. Clone the repository:
git clone https://github.com/Brian5049/AuditSummary.git
2. Navigate to the project directory:
cd project-directory
3. Create a virtual environment:
   # I don't know how much of this you'll need. I didn't need it all. It might be helpful to you though.
- Windows:
  ```
  python -m venv venv
  ```
- macOS/Linux:
  ```
  python3 -m venv venv
  ```
4. Activate the virtual environment:
- Windows:
  ```
  .\venv\Scripts\activate
  ```
- macOS/Linux:
  ```
  source venv/bin/activate
  ```
5. Install dependencies:
pip install -r requirements.txt
pip install matplotlib
pip install reportlab

## Usage

1. Activate the virtual environment:
- Windows:
  ```
  .\venv\Scripts\activate
  ```
- macOS/Linux:
  ```
  source venv/bin/activate
  ```
2. Run the script:
python script.py

## Contributing
None. I don't plan on working continuously with this, I just wanted it to be here as a show of work and help to others who may be working on similar projects.

## License

No license pertains to this script.
