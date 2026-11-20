# DataWeave

**From multiple sources. One standardized dataset.**

Open `Excel Column Mapper.html` in a modern browser. Choose the reference workbook (the `Template` worksheet is read dynamically), then add XLS, XLSX, and CSV sources. Each worksheet is processed independently and may be selected or excluded. Review the original values, confirm each worksheet’s mappings, and inspect the standardized preview.

The mapping engine combines paired first and last names into the contact field. It evaluates `Title` values to suggest a professional designation versus a personal honorific, and leaves mixed or unclear cases for manual review. Mapping confidence includes an explanation; priority and conflicts are visible before export.

The consolidated preview includes a calculated Testing & Validation Summary. It reconciles internal source-row IDs, validates source filename and worksheet metadata, compares mapped values field by field, reports row and field accuracy, and lists worksheet results and inspectable mismatches with original row numbers. Processing completion and validation status are tracked separately in browser-local history.

Only the final appended dataset can be downloaded: `Standardized_Master.xlsx` or `Standardized_Master.csv`. Internal row identifiers are not exported. History and mapping profiles stay in this browser; uploaded records are not saved to history or sent to an application server.

Limits: 14 MB per file, 5,000 data rows per worksheet. The app loads the SheetJS parser and Inter font from their CDNs, so an internet connection is required when opening it.

Brand assets are in `DataWeave Brand/`: light and dark horizontal wordmarks, an app icon and favicon. The app includes the same icon inline.
