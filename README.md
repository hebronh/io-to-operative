# IO to Operative

This tool converts media IO Excel files into Operative-ingestible Sales Order templates (currently limited to Spectrum and Effectv), eliminating manual data entry in Operative One.

The goal is to:

* Reduce manual setup time in Operative
* Prevent ingest errors caused by formatting inconsistencies
* Ensure outputs exactly match Operative’s required templates (values, dates, quantities, formatting)
* Support multi-order IOs by generating multiple Operative files when needed

What This Tool Does

* Accepts an IO Excel file (.xlsx or .xls)
* Parses schedule line items, ignoring non-billable rows (e.g. Property = Ampersand)
* Automatically classifies each line as:
  Spectrum or Effectv

* Splits the IO into multiple outputs if more than one order type is present

* Fills the correct Operative template for each order type

* Outputs .xls files that are ingestible directly into Operative (import order feature)

Operative is extremely strict — formatting must match exactly. This tool aims to preserve template structure and replaces only the relevant data.

## Supported Templates

Templates are stored in (provided by Political Team):

/public/templates/

Currently supported:

operative-spectrum-template.xls

operative-effectv-template.xls

Each template:

* Preserves original formatting and structure

* Has default values cleared/replaced with parsed IO data

* Outputs only valid Operative-approved values

Key Rules / Logic

* Rows where Property === "Ampersand" are ignored

* Dates are written as MM/DD/YYYY strings (not JS Date objects)

* Output format is .xls, not .xlsx

* Quantity, Net Unit Cost, and Line Item Name must match Operative expectations exactly

* Line items retain correct row alignment and formatting from the template

# How to Run Locally
1. Install dependencies
npm install

2. Start the dev server
npm run dev

(The app runs locally via Vite.) 

3. Use the app

Open the app in your browser

Upload an IO Excel file (.xlsx or .xls)

The tool will:

* Parse the IO

* Detect order types

* Generate one or more Operative .xls files

Downloaded files are ready for Operative ingestion

# Project Structure (High Level) 
````md
```text
src/
├── App.tsx
│   └── Application entry point.
│       Handles file upload, template selection, and export orchestration.
│
├── converters/
│   ├── parseSourceIo.ts
│   │   └── Parses raw IO Excel files and normalizes line data
│   │       (markets, dates, net, imps, targeting, order type).
│   │
│   ├── fillTemplate.ts
│   │   └── Fills Operative Excel templates with normalized IO data
│   │       while preserving formatting and ingestion requirements.
│   │
│   └── types.ts
│       └── Shared TypeScript types used across parsing and filling logic.
│
├── templates/
│   └── templateConfig.ts
│       └── Maps OrderType (Spectrum, Effectv, etc.)
│           to the correct Operative template and sheet configuration.
│
public/
└── templates/
    ├── operative-spectrum-template.xls
    │   └── Exact Operative Spectrum ingest template (unchanged formatting).
    │
    └── operative-effectv-template.xls
        └── Exact Operative Effectv ingest template (unchanged formatting).
```` 

# Purpose

Operative One is:

* Slow to configure manually

* Extremely sensitive to formatting errors

* Integrated with platforms like Datamax and AND

This tool aims to:

* Standardize ingest logic

* Prevents common ingest failures

* Saves significant manual effort for Ad Ops

