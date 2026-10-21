# V7.4 validation compatibility fix

This release changes validation behavior only; it does not migrate or rewrite task data.

## Schema mappings

- ShowApp / Master Backlog: ID → Task ID, Feature / Issue → Task Title, Requirement / Description → Description.
- Governance / Repository Register: row 4 headers; ID → Task ID, Repository → Task Title, Purpose / Description → Description, Primary Owner → Owner, Status → Status.
- Nano and Klima retain their existing headers and Google Sheets table structures.

## Validation safety

- Typed Google Sheets table columns are detected from Sheets API table metadata.
- Any typed column is skipped rather than receiving cell-range validation.
- Existing equivalent validation is reported as already present/skipped.
- Existing different validation is preserved and reported as skipped.
- No table is converted to an ordinary range.
- No source header is renamed.
- No task value is changed by validation setup.

## After deployment

As Admin, open Users → Google Sheets data validation → Apply / repair validation.
The result table now separates successful, skipped, and failed field operations.
