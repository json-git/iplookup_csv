# What the script does

When you select a column of IPs and run the script:

1. Inserts a new column to the right
2. Labels it “Registrar”
3. Checks each IP:

   * If already seen → reuse cached result
   * If new → call IPinfo API
4. Writes results next to each IP

---

# Script

Paste this into **Automate → New Script**

```typescript
async function main(workbook: ExcelScript.Workbook) {

    const sheet = workbook.getActiveWorksheet();
    const selectedRange = workbook.getSelectedRange();
    const values = selectedRange.getValues() as (string | number | boolean)[][];
    const startRow = selectedRange.getRowIndex();
    const col = selectedRange.getColumnIndex();
    const token = "api token pls";
    // Insert output column
    const insertRange = sheet.getRangeByIndexes(
        0,
        col + 1,
        sheet.getUsedRange()?.getRowCount() ?? 1000,
        1
    );

    insertRange.insert(ExcelScript.InsertShiftDirection.right);

    // Header
    sheet.getCell(startRow, col + 1).setValue("Registrar");

    // Cache for duplicates
    const cache: { [ip: string]: string } = {};

    for (let i = 0; i < values.length; i++) {
        const ip = String(values[i][0] ?? "").trim();
        if (!ip) continue;
        let result: string;
        if (cache[ip] !== undefined) {
            result = cache[ip];
        } else {
            try {
                const response = await fetch(
                    `https://ipinfo.io/${ip}?token=${token}`
                );
                const data = await response.json() as { org?: string };

                result = data.org ?? "Not Found";
            } catch {
                result = "Lookup Failed";
            }
            cache[ip] = result;
        }
        sheet.getCell(startRow + i, col + 1).setValue(result);
    }
}
```

---

# How to use it

## 1. Prepare your data

* Put IPs in a single column
* Example:

```
8.8.8.8
1.1.1.1
8.8.8.8
```

---

## 2. Select the column

Highlight only the IP column (cells with values) or specific cells.

---

## 3. Run script

Go to:

> Automate tab → Run Script → select your script

---

## 4. Output

Excel will:

* Insert a new column to the right
* Fill it with registrar data

Example:

| IP      | Registrar  |
| ------- | ---------- |
| 8.8.8.8 | Google LLC |
| 1.1.1.1 | Cloudflare |
| 8.8.8.8 | Google LLC |

---
