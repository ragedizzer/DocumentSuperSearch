
# Supersearch README

This tool uses PowerShell and will work in highly restricted offices. With its ability to search multiple terms in one go, then provide results for all search terms in a spreadsheet, it is a great tool for auditing and cleaning up large Word Document repositories or SharePoint document libraries. It only searches the local/sync/network drives. It uses Outlook to send emails.

## Contents
- Overview
- Install / Setup
- Using SharePoint libraries (sync locally)
- GUI Usage (Search-Gui.ps1)
- CLI Usage (Search.ps1)
- Output File Details
- Mail Results (Outlook)
- Sleep Prevention
- Release and maintenance
- Troubleshooting

## Overview
This document explains how to install and use the Supersearch tool, including the GUI and key options.
Supersearch scans Word documents for terms in:
- Document body text
- Hyperlink paths
- Metadata fields (Summary, Notes, Tags, Enterprise Keywords, Author)
Results are saved to an Excel file and can be emailed via Outlook.
Supersearch is a PowerShell-based document discovery tool for Word files. It searches body text, hyperlinks, and key metadata fields, then exports results to a structured Excel report. This makes it ideal for governance and audit workflows where documents must be grouped or validated by metadata such as Summary, Tags, or Enterprise Keywords.

The output file is audit-friendly because it:
- Captures where each match was found (body, metadata field, or link)
- Normalizes metadata into consistent columns for easy filtering/grouping
- Enables quick review of compliance or inventory gaps across large libraries

## Install / Setup
1. Copy the folder Supersearch from Tools and Scripts to your desired location.
2. Click Start and search for "PowerShell ISE" and open it.
3. Click the Open icon, browse to Search-Gui.ps1, and open the file.
4. Press the green triangle play button.
 
## Using SharePoint libraries (sync locally)
This tool searches local files, so SharePoint libraries must be synced to your computer first.
Steps (Microsoft 365 / OneDrive):
1. Open the SharePoint document library in your browser.
2. Click Sync in the command bar.
3. Approve opening OneDrive (if prompted).
4. Wait for the library to appear in File Explorer under your organization's name.
5. Right-click the folder, choose Properties, and make sure Read-only is checked.

**Tip:** The library path typically looks like:

C:\Users\<you>\Your Org Name\Library Name

## GUI Usage (Search-Gui.ps1)
 
**Path controls**
- Search path: the folder (or file) to scan.
- Output folder: where the Excel report is saved.
Buttons (per path):
- "..." = browse
- "S" = save the current field as the default
- "R" = reset to Documents
- This will generate a Search-Gui-Settings.json where this setting is stored.

**Search options**
- Include subfolders
- Search text content
- Search link paths
- Link mode (only visible when Search link paths is checked)
- Search metadata fields
- Include metadata columns
- Search file names
- Match case
- Match whole word
- Keep awake while running
- Doc timeout (seconds)

**Email options**
Check Email results and fill in:
- Email to
- Send on behalf of (optional)
Email is sent through Outlook using user’s credentials.
**Output options**

**Output format:**
- Excel: standard .xlsx
- ExcelTable: .xlsx with a formatted table for filtering
- Csv: .csv for easy import into other tools
**Link mode (use cases)**

**Link mode controls how hyperlink targets are searched:**
- AddressOnly: fastest; searches only the hyperlink address/URL
- AddressAndSub: also searches SubAddress (bookmark/anchor)
- All: address + subaddress + display text (most thorough)

**Use cases:**
- Large libraries where link-only searches are slow → use AddressOnly
- SharePoint docs with anchors → use AddressAndSub
- When link display text may contain the term → use All

**Doc timeout**
Global per-document timeout in seconds. If a document takes too long, it is skipped and the search continues.
- Use when Word hangs on large or corrupted documents
- Set to 0 to disable

**Stop button**
Stops the scan immediately and saves the results collected so far.
If email is enabled, the partial results are emailed as well.

## Metadata fields vs metadata columns
- Search metadata fields: allows terms to match Doc-ID, Summary, Notes, Tags, Enterprise Keywords, Author.
- Include metadata columns: adds those metadata values to the output file, even if you are not searching them.
**Use case:**
- Turn on Include metadata columns to build audit reports even if your search terms are only in the body text or links.

## CLI Usage (Search.ps1)
You can run the search directly from PowerShell:
.\Search.ps1

**Or call the function with options:**
Invoke-DocumentSearch `
- Path "C:\Docs" `
- FindTerms @("term1","term2") `
- IncludeSubfolders $true `

## Output File Details
The report is saved with a name based on the search terms and date:
- Search results - term1, term2 - yyyy-MM-dd.xlsx
- If CSV format is selected, the extension changes to: Search results - term1, term2 - yyyy-MM-dd.csv

**Column order (metadata on)**
1.	MatchedTerms
2.	Found
3.	Doc-ID
4.	FileName
5.	FullPath
6.	Author
7.	Notes
8.	Summary
9.	Tags
10.	Enterprise Keywords

**"Found" column**
Shows where the match occurred:
Body; Link; Doc-ID; Summary; Notes; Tags; Enterprise Keywords; Author
 
**Multiple terms and grouping**
You can search for multiple terms at once (comma or semicolon separated).
The output includes:
- MatchedTerms: which terms were found in the document
- Found: where they were found (Body, Link, Doc-ID, Summary, etc.)

This makes it easy to:
- Filter/group by MatchedTerms to see which term matched
- Filter/group by Found to see whether terms appear in body text vs metadata

## Email Results (Outlook)
Email is sent via Outlook (no SMTP needed).
Set in GUI:
- Email results: checked
- Email to: one or more recipients
- Send on behalf of: optional

## Sleep Prevention
To keep the machine awake during a long scan:
- GUI: "Keep awake while running" (checked by default)
- CLI: -PreventSleep $true

## Release and maintenance
See RELEASING.md for the release checklist and branch protection expectations.

## Troubleshooting
- Output file not updating
Close the Excel report before running. If the file is open, the script will stop and ask you to close it.

- GUI changes not showing
Close and re-open the GUI. It loads Search.ps1 on startup.

- Metadata not found
Some documents do not contain all metadata fields. Empty fields are expected.
