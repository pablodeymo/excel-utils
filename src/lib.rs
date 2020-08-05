use chrono::NaiveDate;

pub fn convert_date(field_value: Option<&&calamine::DataType>) -> Option<NaiveDate> {
    use std::str::FromStr;
    match field_value? {
        calamine::DataType::String(s) => {
            let v: Vec<&str> = s.split("/").collect();
            let day = u32::from_str(v.get(0)?).ok()?;
            let month = u32::from_str(v.get(1)?).ok()?;
            let year = i32::from_str(v.get(2)?).ok()?;
            Some(NaiveDate::from_ymd(year, month, day))
        }
        _ => None,
    }
}

pub fn convert_string(field_value: Option<&&calamine::DataType>) -> Option<String> {
    match field_value {
        Some(calamine::DataType::String(s)) => Some(s.clone()),
        _ => None,
    }
}


#[cfg(test)]
mod tests {
    #[test]
    fn it_works() {
        assert_eq!(2 + 2, 4);
    }
}
