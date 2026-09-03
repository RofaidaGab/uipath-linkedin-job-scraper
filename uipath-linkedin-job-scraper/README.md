# UiPath LinkedIn Job Scraper

A UiPath workflow that automates browsing LinkedIn's job search results, paginating through listings, extracting each job's **title, company, description, and link**, and exporting the collected data to an Excel workbook.

## What it does

`Main.xaml` runs as a single sequence:

1. **`Build Data Table`** — defines the output schema up front: `Job ID`, `Title`, `Company`, `Description`, `Link`.
2. **Chrome Application Card** — attaches to a Chrome session already on a LinkedIn job search results page.
3. **`While` loop (pagination):**
   - **`Extract Table Data`** (UI Automation "Extract Data Generic") pulls the visible **Title** and **Company** for every job card on the current page into a temporary table.
   - Two `Add Data Column` steps extend that table with `Job ID` and `Description` placeholders.
   - **`For Each UI Element`** walks each job card on the page:
     - Logs entry, then inside a `Try/Catch`:
       - Clicks the job card (`Click Current Job`), reads the job's URL (`Get URL`) to extract a numeric **Job ID**.
       - Checks for a "show more" (`daha fazla`) button on the description panel and clicks it if present, so the full description is expanded before reading it.
       - **`Get Description`** reads the expanded job description text.
       - A `Multiple Assign` step cleans/trims the extracted text.
       - Adds a completed row (Job ID, Title, Company, Description, Link) to the results table.
       - On error, the item is logged and skipped (`Log Error - Skip Item`) rather than stopping the whole run.
   - Checks whether a "Next" (`İleri`) pagination control exists; if so, clicks it and continues the `While` loop, otherwise sets the loop's continuation flag to `False` and exits.
4. **`Excel Process Scope` → `Use Excel File`** — opens the output workbook, clears the `Jobs` sheet/table, and writes the full collected `DataTable` in one `Write Range` operation.

## URL parameter encoding & dynamic web extraction

- The automation is designed to be started from **any LinkedIn job search results URL** (e.g. `https://www.linkedin.com/jobs/search-results/?keywords=<term>&start=<offset>`), and works by reading the page's live DOM rather than calling an API — LinkedIn does not expose a public jobs-search API, so this project uses UI Automation (browser-driven scraping) with the modern **UiPath UI Automation "Next" (NApplicationCard / NExtractDataGeneric)** activities.
- Pagination is handled two ways in the workflow:
  - the `start=` query-string offset visible in the recorded browser history increases by a fixed page size as LinkedIn's own "load more" mechanism advances,
  - and/or via clicking LinkedIn's in-page "Next"/"daha fazla" (more) controls directly, so the scraper follows whatever pagination LinkedIn serves rather than manually constructing every page URL.
- Individual job links are built from the numeric Job ID captured per card: `"https://www.linkedin.com/jobs/view/" & str_JobID & "/"`, which is LinkedIn's canonical per-posting URL pattern.
- Element matching uses UiPath's **Object Repository / selector + fuzzy-selector** definitions recorded against the page's DOM structure (`div` nesting paths), so the scraper is resilient to LinkedIn showing slightly different job cards, but **will need re-recording if LinkedIn changes its page layout**.

## DataTable manipulation

- `BuildDataTable` defines a strongly-typed schema up front (`Job ID`, `Title`, `Company`, `Description`, `Link`) so downstream `Write Range` calls are predictable.
- Extracted "generic" table data (Title/Company per page) is merged, column-by-column, into the final results `DataTable` per job card, rather than trying to extract every field in one pass — this keeps the extraction resilient (a failure reading one field doesn't lose the whole row) and lets the `Try/Catch` skip only the problem row.
- The table is cleared and fully rewritten (`Clear Sheet/Range/Table` → `Write Range`) at the end of each run rather than appended to, so re-running the automation always reflects the latest search results.

## Excel output export

- Output goes to **`Data/Output/LinkedIn_RPA_Jobs.xlsx`**, sheet `Jobs`, via `Excel Process Scope` → `Use Excel File` → `Clear Sheet/Range/Table` → `Write Range`.
- The path is relative to the project root, so it works out of the box after cloning as long as the file exists at that location (see **Setup**).

## Setup

1. Clone the repo and open `project.json` in **UiPath Studio** (authored on Studio `26.0.199.0`).
2. Let Studio restore the dependencies in `project.json`: `HtmlAgilityPack`, `UiPath.Excel.Activities`, `UiPath.MicrosoftOffice365.Activities`, `UiPath.System.Activities`, `UiPath.UIAutomation.Activities`, `UiPath.WebAPI.Activities`.
3. Create the output file the workflow expects:
   - Put a workbook named **`LinkedIn_RPA_Jobs.xlsx`** with a sheet named **`Jobs`** inside a `Data/Output/` folder at the project root (`Data/Output/LinkedIn_RPA_Jobs.xlsx`). An empty workbook with just the `Jobs` sheet is enough — the workflow clears and rewrites it on each run.
4. Have **Chrome** installed and be logged in to LinkedIn in that Chrome profile (the automation attaches to an existing browser session rather than logging in itself).
5. Navigate Chrome to a LinkedIn job search results page (e.g. search for a role/keyword you're interested in) before starting the workflow, since the `Chrome Application Card` attaches to whatever job search page is currently open.

## Running the workflow

1. Open the project in Studio, open `Main.xaml`.
2. Make sure Chrome is open on a LinkedIn jobs search-results page (see step 5 above).
3. Press **F5** / **Run**.
4. The robot will page through the results, extracting each job until no "Next"/"daha fazla" control is found, then write everything to `Data/Output/LinkedIn_RPA_Jobs.xlsx`.
5. Check the **Output** panel for `"Total jobs collected: N"` and `"Added to excel"` log messages to confirm a clean run; any skipped items are logged individually rather than stopping the run.

## Notes on scope & fair use

This project is for personal, educational RPA practice (learning UI Automation, data extraction, and Excel export). LinkedIn's Terms of Service restrict automated scraping of the platform — use responsibly, at low volume, on your own account/session, and do not redistribute scraped personal data from other users' profiles or postings.

## Security notes

- No LinkedIn credentials are stored or entered by this workflow — it attaches to an already-authenticated Chrome session, so your LinkedIn password never appears in the project.
- The workbook output path has been changed from a bare filename to `Data/Output/LinkedIn_RPA_Jobs.xlsx` so no personal file-system paths are published in this repo.
