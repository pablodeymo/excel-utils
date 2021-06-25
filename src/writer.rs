use anyhow::{anyhow, Result};
use calamine::DataType;
use chrono::Datelike;
use xlsxwriter::{DateTime, FormatAlignment, FormatColor, Workbook, Worksheet};

pub fn write_header(
    workbook: &Workbook,
    worksheet: &mut Worksheet,
    starting_row: u32,
    fontcolor: FormatColor,
    bgcolor: FormatColor,
    header_titles: &[(&str, f64)],
) -> Result<()> {
    let format = workbook
        .add_format()
        .set_font_color(fontcolor)
        .set_bg_color(bgcolor)
        .set_align(FormatAlignment::CenterAcross);

    for (i, (text, width)) in header_titles.iter().enumerate() {
        let col: xlsxwriter::WorksheetCol = i as u16;
        worksheet
            .set_column(col, col, *width, None)
            .map_err(|e| anyhow!("Error setting width. {:?}", e))?;

        worksheet
            .write_string(starting_row, col, *text, Some(&format))
            .map_err(|e| anyhow!("Error writting header. {:?}", e))?;
    }

    Ok(())
}

pub fn write_content_table(
    workbook: &Workbook,
    worksheet: &mut Worksheet,
    starting_row: u32,
    content_table: &[Vec<Option<DataType>>],
    include_total_row: bool,
    totals: Option<&[(DataType, u16, &str)]>,
) -> Result<()> {
    let datetime_format = workbook.add_format().set_num_format("yyyy-mm-dd");
    let count_rows = content_table.len();

    for (i, row_content) in content_table.iter().enumerate() {
        let row: xlsxwriter::WorksheetRow = i as u32 + starting_row;
        for (j, text) in row_content.iter().enumerate() {
            let col: xlsxwriter::WorksheetCol = j as u16;

            match text {
                Some(DataType::String(s)) => {
                    worksheet
                        .write_string(row, col, s, None)
                        .map_err(|e| anyhow!("Error writting string. {:?}", e))?;
                }
                Some(DataType::Int(v)) => {
                    worksheet
                        .write_number(row, col, *v as f64, None)
                        .map_err(|e| anyhow!("Error writting number. {:?}", e))?;
                }
                Some(DataType::DateTime(d)) => {
                    let chrono_date = chrono::NaiveDate::from_ymd(1899, 12, 30)
                        + chrono::Duration::days((*d).round() as i64);
                    let datetime = DateTime::new(
                        chrono_date.year() as i16,
                        chrono_date.month() as i8,
                        chrono_date.day() as i8,
                        12,
                        0,
                        0.0,
                    );
                    worksheet
                        .write_datetime(row, col, &datetime, Some(&datetime_format))
                        .map_err(|e| anyhow!("Error writting date. {:?}", e))?;
                }
                Some(DataType::Float(f)) => {
                    worksheet
                        .write_number(row, col, *f as f64, None)
                        .map_err(|e| anyhow!("Error writting float. {:?}", e))?;
                }
                _ => {}
            };
        }
    }

    if include_total_row {
        let row: xlsxwriter::WorksheetRow = count_rows as u32 + starting_row;
        let format_total_row = workbook
            .add_format()
            .set_bold()
            .set_border_top(xlsxwriter::FormatBorder::Medium);

        // Print "Total" as a title in the last row
        worksheet
            .write_string(row, 0, "Total", Some(&format_total_row))
            .map_err(|e| anyhow!("Error writting string. {:?}", e))?;

        // Print each formula
        if let Some(totals_val) = totals {
            for (total_value, totals_col, col_letter) in totals_val {
                let col: xlsxwriter::WorksheetCol = *totals_col;
                let value: f64 = match *total_value {
                    DataType::Float(f) => f,
                    DataType::Int(v) => v as f64,
                    _ => 0.0,
                };
                let formula = format!(
                    "=SUM({}{}:{}{})",
                    col_letter,
                    starting_row + 1,
                    col_letter,
                    count_rows as u32 + starting_row
                );
                worksheet
                    .write_formula_num(row, col, &formula, Some(&format_total_row), value)
                    .map_err(|e| anyhow!("Error write formula num. {:?}", e))?;
            }
        }
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
pub fn write_table(
    workbook: &Workbook,
    worksheet: &mut Worksheet,
    starting_row: u32,
    header_fontcolor: FormatColor,
    header_bgcolor: FormatColor,
    header_titles: &[(&str, f64)],
    content_table: &[Vec<Option<DataType>>],
    include_total_row: bool,
    totals: Option<&[(DataType, u16, &str)]>,
) -> Result<()> {
    // Write the header
    write_header(
        workbook,
        worksheet,
        starting_row,
        header_fontcolor,
        header_bgcolor,
        header_titles,
    )?;
    write_content_table(
        workbook,
        worksheet,
        starting_row + 1,
        content_table,
        include_total_row,
        totals,
    )
}

#[cfg(test)]
mod tests {
    use calamine::DataType;
    use xlsxwriter::{FormatColor, Workbook};

    #[test]
    fn write_header_test() {
        let header = [
            ("Date", 10.0),
            ("Count", 5.0),
            ("Description", 15.0),
            ("Amount", 12.0),
        ];

        let workbook = Workbook::new("tests/header1.xlsx");
        let mut sheet1 = workbook.add_worksheet(None).unwrap();
        let content: Vec<Vec<Option<DataType>>> = vec![
            vec![
                Some(DataType::DateTime(44289.0)),
                Some(DataType::Int(10)),
                Some(DataType::String("Pencil".to_string())),
                Some(DataType::Float(1.35)),
            ],
            vec![
                Some(DataType::DateTime(44288.0)),
                Some(DataType::Int(5)),
                Some(DataType::String("Notepad".to_string())),
                Some(DataType::Float(5.70)),
            ],
        ];
        let totals: [(DataType, u16, &str); 1] = [(DataType::Float(7.05), 3_u16, &"D")];

        crate::writer::write_table(
            &workbook,
            &mut sheet1,
            2,
            FormatColor::White,
            FormatColor::Navy,
            &header,
            &content,
            true,
            Some(&totals),
        )
        .unwrap();
        workbook.close().unwrap();
    }
}
