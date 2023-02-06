use anyhow::{anyhow, Result};
use calamine::{open_workbook_auto, DataType, Range, Reader};
use std::path::Path;

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

#[cfg(test)]
mod tests {
    use calamine::DataType;
    use pretty_assertions::assert_eq;

    #[test]
    fn open_file() {
        let range = super::open_nth_workbook_from_file("src/test/open.xls", 0).unwrap();
        assert_eq!(range.get_value((0, 0)), Some(&DataType::Int(1)));
    }
}
