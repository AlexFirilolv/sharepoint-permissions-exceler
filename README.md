# SharePoint Permissions Excel Reporter

Tired of executives and stakeholders asking you to prepare a report of your SharePoint site's permissions "in a readable way",
but the multi-billion company running this mess of an intranet service can't provide a built-in way to do so,
besides a pesky CSV export masquerading itself as a report for items shared with external users?

Can't or plainly won't accept to pay for a third party tool that does that, like ShareGate?

Well, this tool converts the POS SharePoint CSV permission exports into decently readable and formatted Excel reports with Entra ID security group member expansion.

## What It Does

Takes a SharePoint permissions CSV export and generates two formatted Excel reports:

- **Parent_Permissions.xlsx** - Permissions for parent-level resources (Web and List items)
- **Unique_Permissions.xlsx** - Items with unique permissions, grouped by document library

Both reports include:
- Formatted Excel tables with auto-adjusted columns
- Expanded security group membership details (when Entra ID credentials are provided)

## Requirements

- Python 3.7+
- SharePoint Online permissions export CSV file (the lousy CSV permissions report from SharePoint Online)

## Getting the SharePoint Permissions CSV

1. Navigate to your SharePoint Online site
2. Click the **Settings icon** (gear icon) in the top navigation bar
3. Select **Site Usage**
4. Scroll to the bottom of the 'Site Usage' page
5. Click **Run report**
6. Choose where to save the generated permission report
7. Download the report and either:
   - Save it in the directory of this cloned repo (it will be auto-detected if the filename starts with `Files_`), or
   - Save it anywhere and provide the full path interactively when running the tool via the terminal

## Installation

```bash
# Clone the repository
git clone https://github.com/AlexFirilolv/sharepoint-permissions-exceler.git
cd sharepoint-permissions-exceler

# Create virtual environment (isolates dependencies from system Python)
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

## Usage

### Basic Usage (Without Security Group Expansion)

```bash
python main.py
```

The script will:
1. Auto-detect CSV files starting with `Files_` or prompt you for the file path
2. Ask if you want to configure Entra ID credentials (or use credentials from `.env`)
3. Generate both Excel reports

### With Entra ID Integration (Recommended)

To expand security groups and show individual members:

1. **Create Entra ID App Registration**:
   - Go to [Azure Portal - App Registrations](https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps)
   - Create new registration
   - Note the **Tenant ID** and **Client ID**

2. **Add API Permissions** (Application permissions, not Delegated):
   - `GroupMember.Read.All`
   - `User.Read.All`
   - Grant admin consent

3. **Create Client Secret**:
   - Go to "Certificates & secrets"
   - Create new secret and copy the value

4. **Configure credentials** (optional - you can also enter them when running the tool):
   ```bash
   cp .env.example .env
   # Edit .env with your credentials, or skip this and provide them interactively
   ```

5. **Run the script**:
   ```bash
   python main.py
   ```

   If no `.env` file is configured, the tool will prompt you to enter credentials in the terminal.

## Output

Each Excel file contains:
- Main data table with permissions
- Security group members section (if Entra ID is configured) showing:
  - Display Name
  - Email
  - Member Type

## License

MIT
