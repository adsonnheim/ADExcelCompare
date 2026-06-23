# ADExcelCompare
PowerShell script to check the devices in both an excel spreadsheet and a domain OU.

How to use:
- Create a copy of `.env.example` and rename it to `.env.local`
- Fill out parameters in newly created .env.local file of where your users are located in your domain, e.g.:
    - `SEARCH_BASE=OU=users,DC=domain,DC=my`
    - `TRACKED_SHEETS=Sheet 1, Sheet 2, Sheet 3`
    - `REPORT_URL=/path/to/where/file/is/located`
- Run `ADExcelCompare.ps1`
- View results in terminal