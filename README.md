rust-wasm binding that uses calamine to read XLSX files

```typescript
import {BasicXLSXReader} from "./basic_xlsx_reader";

const readFileAsUInt8Array = (file: File): Int8Array => {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function () {
            const data = reader.result;
            const array = new Int8Array(data);
            resolve(array);
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
};

const input = document.querySelector('input[type=file]');
const fileUint8Array = await readFileAsUInt8Array([input.files]);
const basicXLSXReader = new BasicXLSXReader(fileUint8Array);
const output = basicXLSXReader.read({
    headerRow: 1, // 0 indexed
    sheet: {name: "Test Sheet"},
    includeEmptyCells: true,
});

console.log(output);

/*
{
  "sheets": [
    {
      "sheet": "Test Sheet",
      "row": [
        {
          "columns": [
            {
                "header": "Column 1",
                "value": "Test Cell"
            },
            {
                "header": "Column 2",
                "value": 2
            }
          ]
        }
      ]
    }
  ]
}

 */

```