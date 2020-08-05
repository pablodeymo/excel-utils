use chrono::NaiveDate;

pub fn convert_date(field_value: Option<&&calamine::DataType>) -> Option<NaiveDate> {
    use std::str::FromStr;
    match field_value? {
        calamine::DataType::String(s) => {
            // if the fields are not separated by /, they are all together
            let v: Vec<&str> = s.split("/").collect();
            if v.is_empty() {
                let day = u32::from_str(v.get(0)?).ok()?;
                let month = u32::from_str(v.get(1)?).ok()?;
                let year = i32::from_str(v.get(2)?).ok()?;
                Some(NaiveDate::from_ymd(year, month, day))
            } else if s.len() == 8 {
                let day = u32::from_str(s.get(0..2)?).ok()?;
                let month = u32::from_str(s.get(2..4)?).ok()?;
                let year = i32::from_str(s.get(4..8)?).ok()?;
                Some(NaiveDate::from_ymd(year, month, day))
            } else {
                None
            }
        }
        _ => None,
    }
}

pub fn convert_string(field_value: Option<&&calamine::DataType>) -> Option<String> {
    match field_value {
        Some(calamine::DataType::String(s)) => Some(s.clone()),
        Some(calamine::DataType::Float(f)) => Some(format!("{}", f)),
        _ => None,
    }
}

pub fn convert_i32(field_value: Option<&&calamine::DataType>) -> Option<i32> {
    match field_value {
        Some(calamine::DataType::Float(f)) => Some(*f as i32),
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
