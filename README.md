# BCSL Outstanding-Email Generator (RPA Tool)

This tool reads a Tally **"Sundry Debtors / Pending Bills"** export spreadsheet (`Client_OS_Report.xlsx`) and automates email dispatch:
1. **Client Job:** Sends each client a table of their outstanding invoices split into two ageing sections (older or newer than the threshold days).
2. **Sales Job:** Sends each sales manager a single consolidated list of all outstanding invoices strictly older than the threshold days for all of their clients.

---

## 🛠️ Step 1: Prepare and Set Up the Python Environment (`.py` setup)

Before running the Python script directly, make sure Python and the required libraries are installed.

### 1. Install Python
Download and install Python (v3.8 or higher) on your machine from [python.org](https://www.python.org/). Make sure to check **"Add Python to PATH"** during installation.

### 2. Install Required Libraries
Open your command prompt or PowerShell and run the following command to install the required Excel parsing and Windows integration libraries:
```powershell
pip install openpyxl pywin32
```
*Note: `pywin32` is only required if you choose to send emails via the locally-installed Outlook desktop app (`TRANSPORT=outlook`).*

### 3. File Setup
Place the following files together in your project folder:
* `bcsl_outstanding_emailer.py` (the Python script)
* `Client_OS_Report.xlsx` (your Tally outstanding report)
* `bcsl-client-outstanding-email-addresses.xlsx` (client address directory)
* `bcsl-salesteam-email-addresses.xlsx` (sales team address directory)
* `assets/` (optional folder containing logo and social icon images)

### 4. Run the Python Script
You can execute the script using the following command:
```powershell
python bcsl_outstanding_emailer.py
```
On first run, it will automatically generate a default `.env` configuration file in the same folder.

---

## 📦 Step 2: Build the Standalone Executable (`RPA.exe`)

You can compile this Python script into a standalone `.exe` so it can be run on computers that do not have Python installed.

### 1. Install PyInstaller
Install the PyInstaller compiler library:
```powershell
pip install pyinstaller
```

### 2. Build the Executable
Run the following build command in your terminal from the project directory:
```powershell
pyinstaller --onefile --name RPA --console bcsl_outstanding_emailer.py
```
* **`--onefile`**: Bundles all libraries and the python script into a single executable file.
* **`--name RPA`**: Names the resulting file `RPA.exe`.
* **`--console`**: Keeps the command prompt window open when running so your team can view runtime logs, skips, and errors.

### 3. Retrieve the Executable
Once the compilation completes, the standalone `RPA.exe` will be located in the newly created **`dist/`** directory. Move `RPA.exe` to a permanent location along with your spreadsheet files, `.env` file, and `assets/` folder.

---

## ⚙️ Step 3: Configure Parameters in the `.env` File

The `.env` file allows you to customize credentials, paths, limits, and signature details without altering the code or rebuilding the `.exe` file.

| Field Name | Type | Default Value | Description / Accepted Inputs |
| :--- | :--- | :--- | :--- |
| **`RUN_JOBS`** | String | `both` | Which jobs to run. Valid options: `client` (client statements only), `sales` (sales manager consolidated reports only), `both` (both sequentially). |
| **`REPORT_FILE`** | String | `Client_OS_Report.xlsx` | The filename of the Tally exported spreadsheet. Must sit next to the executable. |
| **`ASSETS_DIR`** | String | `assets` | Folder containing logo and social media icon images. |
| **`OUTPUT_DIR`** | String | `outstanding_emails_output` | Directory where logs, manifests, and drafts/previews are generated. |
| **`TRANSPORT`** | String | `dry_run` | Delivery transport mode. Valid options:<br>• `dry_run` (saves local drafts under output folder for verification)<br>• `smtp` (sends via SMTP server configuration)<br>• `outlook` (sends via Windows Outlook desktop application signed-in profile). |
| **`EMAIL_DELAY_SECONDS`** | Float | `2.0` | Time delay (in seconds) to pause between sending each email. Helps avoid spam filters or rate limiting. Options: any number (e.g. 0 to disable, 2.5, 5, etc.). |
| **`SEND_ONLY_WHEN_FLAG_Y`** | Boolean | `True` | Whether to restrict dispatches to clients/reps marked `Y` or `y` in the address books. Options: `True`, `False`. |
| **`ALLOW_FUZZY_MATCH`** | Boolean | `False` | Whether to allow loose matching of client names if spellings vary. Options: `True`, `False`. |
| **`FUZZY_CUTOFF`** | Float | `0.92` | Fuzzy matching criteria score. Ranges from `0.0` (loose) to `1.0` (exact). |
| **`SMTP_HOST`** | String | `smtp.gmail.com` | SMTP server address (only used if `TRANSPORT=smtp`). |
| **`SMTP_PORT`** | Integer | `587` | SMTP port number (commonly `587` for STARTTLS, `465` for SSL). |
| **`SMTP_SECURITY`** | String | `starttls` | Connection security. Options: `starttls`, `ssl`, `none`. |
| **`SMTP_USERNAME`** | String | `jay.mvbom.plays@gmail.com` | Login username/email for the SMTP server. |
| **`SMTP_PASSWORD`** | String | `sbbu olzi xvcp ktjd` | App Password or password credential. (Use App Password for Gmail/Office 365). |
| **`COMPANY_NAME`** | String | `Benchmark Computer Solutions Limited` | Company name used in signatures and titles. |
| **`COMPANY_ADDR`** | JSON Array | `["Line 1", "Line 2"]` | Array of company address lines (e.g. `["501, 5th Floor...", "Marol Naka..."]`). |
| **`COMPANY_CERT`** | String | `ISO 9001:2015 CERTIFIED` | ISO certification text displayed in signatures. |
| **`SOCIAL_LINKS`** | JSON Array | `[["icon.png", "url"]]` | Social icon image file and account link pairs. |
| **`LOGO_FILE`** | String | `image001.png` | Primary signature logo filename inside the assets directory. |
| **`CLIENT_ADDRESS_FILE`**| String | `bcsl-client-outstanding-email-addresses.xlsx` | Client address book spreadsheet name. |
| **`CLIENT_FROM_NAME`** | String | `Accounts - Benchmark Computer Solutions` | From name header on client outstanding emails. |
| **`CLIENT_FROM_EMAIL`** | String | `accounts@benchmarksolution.com` | From email header on client outstanding emails. |
| **`CLIENT_SUBJECT`** | String | `Outstanding - {name}` | Client email subject line template. `{name}` inserts the client's company name. |
| **`CLIENT_THRESHOLD_DAYS`**| Integer | `30` | Overdue threshold split criteria in client statements. |
| **`CLIENT_GREETING`** | String | `Dear Sir/Mam,` | Email opening greeting. |
| **`CLIENT_INTRO_OVER_30`**| String | `Please Update Payment Status,` | Text above the table of invoices older than threshold days. |
| **`CLIENT_INTRO_UPTO_30`**| String | `Please confirm below invoice...` | Text above the table of invoices newer/equal to threshold days. |
| **`CLIENT_SIGNER`** | JSON Dict | `{"name": "...", "title": "..."}` | Contact card dictionary for the client email signer. |
| **`SALES_ADDRESS_FILE`** | String | `bcsl-salesteam-email-addresses.xlsx` | Sales manager address book spreadsheet name. |
| **`SALES_FROM_NAME`** | String | `Mohit P - Benchmark Computer Solutions` | From name header on sales consolidated emails. |
| **`SALES_FROM_EMAIL`** | String | `mohit.p@benchmarksolution.com` | From email header on sales consolidated emails. |
| **`SALES_SUBJECT`** | String | `Outstanding Bills more than 35 days` | Consolidated sales email subject line. |
| **`SALES_THRESHOLD_DAYS`**| Integer | `35` | Filter criteria (only displays invoices strictly older than this number of days). |
| **`SALES_GREETING`** | String | `Dear Sir,` | Sales email opening greeting. |
| **`SALES_INTRO`** | String | `Please find below outstanding bills...` | Explanatory intro paragraph on sales emails. |
| **`SALES_SIGNER`** | JSON Dict | `{"name": "...", "title": "..."}` | Contact card dictionary for the sales email signer. |
