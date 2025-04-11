MOT Checker Tool
A Python utility for automatically checking MOT (Ministry of Transport) test due dates for UK vehicles by querying the DVSA API.
Overview
This tool allows you to:

Batch process vehicle registration numbers from Excel files
Retrieve MOT expiry dates from the official DVSA API
Update your Excel files with the retrieved information
Handle multiple sheets in a single Excel workbook
Process new vehicles with calculated due dates (3 years from registration)

Features

Bulk Processing: Process hundreds of vehicles in a single run
Multi-Sheet Support: Processes all sheets in your Excel workbook
Smart Detection: Automatically identifies columns containing registration numbers
Token Management: Efficiently handles API token caching and renewal
Detailed Logging: Comprehensive logs for troubleshooting
Rate Limiting: Built-in delays to comply with API usage policies

Requirements

Python 3.6+
Required Python packages:

requests
pandas
openpyxl



Installation

Clone this repository:

bash
