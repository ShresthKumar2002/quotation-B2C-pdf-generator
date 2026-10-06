# QuotationPro — B2C Mastersheet edition

## Google Sheet
Set `GOOGLE_SHEET_URL` to the B2C Mastersheet spreadsheet URL. The application reads these tabs:

- `Main`: B = Item Description, C = Vol, D = Rate, E = flags/features. Data starts at row 2.
- `Size`: reserved for the current master-data structure.
- `Vehicle selection`: A = vehicle_id, C = minimum capacity, D = maximum capacity.
- `Vehicle data`: A = vehicle_id, B = city, D = fixed vehicle cost, E = base manpower, F = manpower cost/person.
- `Add ons`: A = flag, B = add-on description, C = cost per item.
- `Packaging charge`: B2 = standard packaging charge, B3 = fragile packaging charge.

The spreadsheet must be accessible to the server. A Google Sheet URL alone is not an API credential; for the current implementation the sheet is fetched through the Google visualization CSV endpoint, so the sheet must be shared/published in a way that permits that access.

## Item rules
Only descriptions are shown in the item search. Quantity is entered manually. Backend re-fetches the selected item from `Main` and does not trust client-supplied rate/volume/flags.

- `item.amount = Rate * qty`
- `item.volume = Vol * qty`
- `vol_of_items = sum(item.volume)`
- flags are parsed from Main column E; recognized flags are 14 through 23.
- Security deposit = net item amount/subtotal.
- Net GST = 18% of net item amount.
- Net monthly storage = net item amount + net GST.

## Logistics rules
1. Find a vehicle whose Vehicle selection capacity range contains `vol_of_items`.
2. Match its vehicle_id + selected city in Vehicle data.
3. If flag 15 is triggered and base manpower < 4, base manpower becomes 4.
4. Packing charge is recalculated from selected items:
   - flag 14 -> volume × fragile packaging charge
   - otherwise -> volume × standard packaging charge
5. Logistics = vehicle fixed cost + base manpower × manpower cost/person + packing charge.

## Add-on rules
For every selected item and every triggered flag 16–23, find the matching row in Add ons and create:
`<item description> - <add-on description>` with the selected item's quantity and `cost per item × qty`.

Token booking amount remains manually entered.

## Run
```powershell
pip install -r requirements.txt
$env:GOOGLE_SHEET_URL="https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/edit"
uvicorn app.main:app --reload --reload-dir app
```


### Faster development updates

Use `--reload-dir app` so WatchFiles monitors only the application folder instead of the entire project directory. This avoids unnecessary reload scanning when generated PDFs, temporary files, or other folders change.

The frontend also debounces quotation previews by 250 ms, so changing quantity does not send a Google-Sheets-backed request on every keystroke.
