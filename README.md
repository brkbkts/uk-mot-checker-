# MOT Checker Tool

A Python utility for automatically retrieving and updating MOT (Ministry of Transport) test expiry dates for UK vehicles directly in Excel spreadsheets.

## Overview

This tool connects to the DVSA (Driver and Vehicle Standards Agency) API to fetch MOT information for vehicle registration numbers stored in Excel files. It processes all sheets in the workbook, finds columns containing registration numbers, and adds or updates an "MOT Due" column with the expiry date information.

## Features

- Bulk processes vehicle registrations across multiple Excel worksheets
- OAuth authentication with token caching to minimize API requests
- Smart detection of registration numbers
- Handles different vehicle scenarios:
  - Vehicles with existing MOT tests
  - Newly registered vehicles with future MOT test dates
  - Very new vehicles (calculating due date as 3 years from registration)
- Preserves original Excel file and creates a new timestamped output file
- Detailed logging for troubleshooting

## Prerequisites

- Python 3.6+
- DVSA API credentials (client ID, client secret, and API key)
- Required Python packages (see Requirements)

## Requirements

The following Python packages are required:
```
requests
pandas
openpyxl
```

You can install them via pip:
```bash
pip install requests pandas openpyxl
```

## Configuration

Before using the tool, you need to add your DVSA API credentials to the script:

1. Open the script and locate the API Configuration section
2. Replace the empty values with your actual credentials:
   ```python
   CLIENT_ID = "your_client_id_here"
   CLIENT_SECRET = "your_client_secret_here"
   API_KEY = "your_api_key_here"
   ```
3. You may also need to update the TOKEN_URL if provided by DVSA

## Usage

### Command Line

Run the script from the command line:

```bash
python mot_checker.py [excel_file_path]
```

If you don't provide the Excel file path as an argument, the script will prompt you to enter it.

### Input Format

The script expects:
- An Excel file (.xlsx) with one or more worksheets
- The first column of each worksheet should contain vehicle registration numbers
- If an "MOT Due" column exists, it will be updated; otherwise, it will be created

### Output

The script creates:
- A new Excel file with "_MOT_Updated_[timestamp]" appended to the original filename
- A log file (mot_checker.log) with detailed processing information

## Example Output

For each vehicle, the MOT Due column will contain one of these:
- A date in DD/MM/YYYY format (for vehicles with MOT expiry dates)
- "New - Due: DD/MM/YYYY" (for new vehicles, calculating 3 years from registration)
- "Vehicle not found" (if the registration is not found in the database)
- Various error messages explaining any issues encountered

## Rate Limiting

The script includes a 1-second delay between API requests to avoid overwhelming the DVSA API server. This can be adjusted if needed, but be mindful of any rate limits imposed by the API.

## Error Handling

The tool includes comprehensive error handling:
- Invalid registration numbers are skipped
- API connection issues are logged
- Missing or inaccessible Excel files trigger appropriate error messages

## License

MIT

## Disclaimer

This tool is provided as-is with no warranty. Users are responsible for ensuring they have appropriate access to the DVSA API and for complying with all relevant terms of service.

## Getting DVSA API Credentials

To obtain the necessary API credentials:
1. Visit the [DVSA API Portal](https://dvsa.api.gov.uk/)
2. Register for an account and request access to the MOT history API
3. Once approved, you'll receive the client ID, client secret, and API key needed for this tool

## Support

For issues or questions, please [open an issue]
