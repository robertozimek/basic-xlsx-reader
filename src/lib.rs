use wasm_bindgen::prelude::*;
use calamine::{Reader, Xlsx, open_workbook_from_rs, HeaderRow, Data, Range};
use js_sys::{Uint8Array};
use std::io::BufReader;
use std::io::Cursor;
use serde::{Deserialize, Serialize};
extern crate console_error_panic_hook;
use std::panic;
use tsify::Tsify;

#[derive(Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi)]
pub struct Result {
    sheets: Vec<SheetResult>,
}

#[derive(Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi)]
pub struct SheetResult {
    pub sheet: String,
    pub rows: Vec<CellResult>,
}

#[derive(Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi)]
pub struct CellResult {
    pub columns: Vec<CellHeaderValue>,
}

#[derive(Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi)]
pub struct CellHeaderValue {
    header: String,
    value: CellValue,
}

#[derive(Serialize, Deserialize, Tsify)]
#[serde(untagged)]
#[tsify(into_wasm_abi)]
pub enum CellValue {
    Int(i64),
    Float(f64),
    String(String),
    Bool(bool),
    Empty,
}

#[derive(Serialize, Deserialize, Tsify)]
#[serde(rename_all = "camelCase")]
#[tsify(from_wasm_abi)]
pub enum SheetOption {
    Name(String),
    Index(usize),
}

#[derive(Serialize, Deserialize, Tsify)]
#[tsify(from_wasm_abi)]
#[serde(rename_all = "camelCase")]
pub struct ReadXLSXOptions {
    #[tsify(optional)]
    header_row: Option<u32>,
    #[tsify(optional)]
    sheet: Option<SheetOption>,
    #[tsify(optional)]
    include_empty_cells: Option<bool>,
}


fn get_rows_for_sheet(sheet_name: &String, sheet_data: &Range<Data>, include_empty: bool) -> SheetResult {
    let mut sheet_body = SheetResult {
        sheet: sheet_name.to_string(),
        rows: vec![],
    };

    let headers = sheet_data.headers().unwrap_or_default();
    let rows = sheet_data.rows();

    for (index, row) in rows.enumerate() {
        // skip header row
        if index == 0 {
            continue;
        }

        let mut result_row = CellResult {
            columns: vec![]
        };
        for (index, cell) in row.iter().enumerate() {
            let cell_value = match cell {
                Data::Empty => if include_empty { CellValue::Empty } else { continue },
                Data::String(s) | Data::DateTimeIso(s) | Data::DurationIso(s) => {
                    CellValue::String(format!("{}", s))
                }
                Data::Float(f) => CellValue::Float(*f),
                Data::DateTime(d) => CellValue::Float(d.as_f64()),
                Data::Int(i) => CellValue::Int(*i),
                Data::Error(e) => CellValue::String(format!("{}", e)),
                Data::Bool(b) => CellValue::Bool(*b),
            };

            result_row.columns.push(CellHeaderValue {
                header: headers.get(index).unwrap_or(&String::from("")).to_string(),
                value: cell_value,
            });
        }
        sheet_body.rows.push(result_row);
    }
    sheet_body
}

#[wasm_bindgen]
pub struct BasicXLSXReader {
    file_bytes: Vec<u8>,
}

#[wasm_bindgen]
impl BasicXLSXReader {
    #[wasm_bindgen(constructor)]
    pub fn new(byte_array: Uint8Array) -> BasicXLSXReader {
        panic::set_hook(Box::new(console_error_panic_hook::hook));

        Self {
            file_bytes: byte_array.to_vec(),
        }
    }

    pub fn read(&self, options: Option<ReadXLSXOptions>) -> Result {
        let options = options.unwrap_or(ReadXLSXOptions {
            header_row: Some(0),
            sheet: None,
            include_empty_cells: Some(false),
        });

        let cursor = Cursor::new(&self.file_bytes);
        let buffer_reader = BufReader::new(cursor);
        let mut excel: Xlsx<_> = open_workbook_from_rs(buffer_reader)
            .expect("Failed to open workbook");

        excel.with_header_row(HeaderRow::Row(options.header_row.unwrap_or(0)));

        let mut results = Result {
            sheets: vec![]
        };

        let include_empty = options.include_empty_cells.unwrap_or(false);
        match options.sheet {
            Some(sheet) => {
                let sheet_name = match sheet {
                    SheetOption::Name(sheet_name) => {
                        sheet_name.to_string()
                    }
                    SheetOption::Index(selected_index) => {
                        excel
                            .sheet_names()
                            .get(selected_index)
                            .expect("Sheet at index not found")
                            .to_string()
                    }
                };

                let sheet = excel.worksheet_range(sheet_name.as_str()).expect("Sheet not found");
                let sheet_body = get_rows_for_sheet(&sheet_name, &sheet, include_empty);
                results.sheets.push(sheet_body);
            }
            None => {
                for (sheet_name, sheet_data) in excel.worksheets().iter() {
                    let sheet_body = get_rows_for_sheet(&sheet_name, &sheet_data, include_empty);
                    results.sheets.push(sheet_body);
                }
            }
        }

        results
    }
}