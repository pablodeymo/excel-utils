use anyhow::{anyhow, Result};
use bigdecimal::BigDecimal;
use calamine::{open_workbook_auto, DataType, Range, Reader};
use chrono::NaiveDate;
use std::path::Path;
use std::str::FromStr;

#[cfg(feature = "enablewriter")]
pub mod writer;

/// Opens the nth worksheet of an Excel file
///
/// # Params
///
/// * `path` - path of the Excel file.
/// * `worksheet_number` number of sheet to open.
///
/// # Errors
///
/// Return Error in case there is an error reading the file.
pub fn open_nth_workbook_from_file(path: &str, worksheet_number: usize) -> Result<Range<DataType>> {
    Ok(match open_workbook_auto(Path::new(path))? {
        calamine::Sheets::Xls(mut sheet) => sheet
            .worksheet_range_at(worksheet_number)
            .ok_or_else(|| anyhow!("Error reading file"))?
            .map_err(|e| anyhow!("Error {}", e))?,
        calamine::Sheets::Xlsx(mut sheet) => sheet
            .worksheet_range_at(worksheet_number)
            .ok_or_else(|| anyhow!("Error reading file"))?
            .map_err(|e| anyhow!("Error {}", e))?,
        calamine::Sheets::Xlsb(mut sheet) => sheet
            .worksheet_range_at(worksheet_number)
            .ok_or_else(|| anyhow!("Error reading file"))?
            .map_err(|e| anyhow!("Error {}", e))?,
        calamine::Sheets::Ods(mut sheet) => sheet
            .worksheet_range_at(worksheet_number)
            .ok_or_else(|| anyhow!("Error reading file"))?
            .map_err(|e| anyhow!("Error {}", e))?,
    })
}

#[must_use]
pub fn convert_date(field_value: &&DataType) -> Option<NaiveDate> {
    match field_value {
        DataType::String(s) => {
            // if the fields are not separated by /, they are all together
            let v: Vec<&str> = s.split('/').collect();
            if v.len() == 3 {
                let day = u32::from_str(v.get(0)?).ok()?;
                let month = u32::from_str(v.get(1)?).ok()?;
                let year = i32::from_str(v.get(2)?).ok()?;
                NaiveDate::from_ymd_opt(year, month, day)
            } else if v.len() == 1 {
                convert_str_to_date(s)
            } else {
                None
            }
        }
        DataType::Float(f) => {
            let num_str = ((*f).round() as u64).to_string();
            if num_str.len() == 8 || num_str.len() == 7 {
                // if the field is a number, it could be the date in full digits
                convert_str_to_date(&num_str)
            } else {
                // sometimes dates are encoded as floats in Excel
                // value 1899-12-30 + f days
                Some(
                    NaiveDate::from_ymd(1899, 12, 30) + chrono::Duration::days((*f).round() as i64),
                )
            }
        }
        _ => None,
    }
}

#[must_use]
pub fn convert_string(field_value: &&DataType) -> Option<String> {
    match field_value {
        DataType::String(s) => {
            if s.trim().is_empty() {
                None
            } else {
                Some(s.trim().to_string())
            }
        }
        DataType::Float(f) => Some(format!("{}", f)),
        DataType::Int(i) => Some(format!("{}", i)),
        _ => None,
    }
}

#[must_use]
pub fn convert_i32(field_value: &&DataType) -> Option<i32> {
    match field_value {
        DataType::Float(f) => Some((*f).round() as i32),
        DataType::Int(i) => Some(*i as i32),
        _ => None,
    }
}

/// Converts float value in cell to decimal, rounding to 2 decimals
#[must_use]
pub fn convert_decimal(field_value: &&DataType) -> Option<BigDecimal> {
    match field_value {
        DataType::Float(f) => Some(BigDecimal::from((*f * 100_f64).round() as i32) / 100),
        DataType::Int(i) => Some(BigDecimal::from(*i as i32)),
        DataType::String(s) => BigDecimal::from_str(&s.replace(",", ".")).ok(),
        _ => None,
    }
}

#[must_use]
pub fn convert_naivedate_to_datatype(date: NaiveDate) -> DataType {
    let ret = NaiveDate::signed_duration_since(date, NaiveDate::from_ymd(1899, 12, 30));
    DataType::DateTime(ret.num_days() as f64)
}

// Auxiliar functions
fn convert_str_to_date(input: &str) -> Option<NaiveDate> {
    let s = if input.len() == 8 {
        input.to_string()
    } else {
        format!("0{}", input)
    };
    let day = u32::from_str(s.get(0..2)?).ok()?;
    let month = u32::from_str(s.get(2..4)?).ok()?;
    let year = i32::from_str(s.get(4..8)?).ok()?;
    NaiveDate::from_ymd_opt(year, month, day)
}

#[cfg(test)]
mod tests {
    use calamine::DataType;
    use chrono::{Datelike, NaiveDate};
    use pretty_assertions::assert_eq;

    #[test]
    fn float_date() {
        let date_test = DataType::Float(16112020.0);
        let ret = crate::convert_date(&&date_test).unwrap();
        assert_eq!(ret.year(), 2020);
        assert_eq!(ret.month(), 11);
        assert_eq!(ret.day(), 16);
    }

    #[test]
    fn open_file() {
        let range = crate::open_nth_workbook_from_file("src/test/open.xls", 0).unwrap();
        assert_eq!(range.get_value((0, 0)), Some(&DataType::Int(1)));
    }

    #[test]
    fn convert_naivedate_to_datatype_test() {
        let d = NaiveDate::from_ymd(1899, 12, 31);
        let res = crate::convert_naivedate_to_datatype(d);
        if let DataType::DateTime(date) = res {
            assert_eq!(date, 1.0);
        } else {
            panic!();
        }
    }
}
