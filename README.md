# DataWeave

**From multiple sources. One standardized dataset.**

Open `Excel Column Mapper.html` in a modern browser. At the start of a session, choose the supplied reference workbook (the `Template` worksheet is read dynamically), then add XLS, XLSX, and CSV sources. Each worksheet is its own source and can be selected or excluded. Review the original values, edit and confirm each worksheet’s mapping, set priority when multiple columns map to one field, and inspect the consolidated preview.

The app exports only the final appended dataset as `Standardized_Master.xlsx` or `Standardized_Master.csv`. Processing history and mapping profiles stay in the browser’s local storage; uploaded records are not saved to history or sent to an application server.

Limits: 14 MB per file, 5,000 data rows per worksheet. The app uses the SheetJS parser and Inter font from their CDNs, so an internet connection is required when opening it.

Brand assets are in `DataWeave Brand/`: light and dark horizontal wordmarks, an app icon and favicon. The app includes the same icon inline.
