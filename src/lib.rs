use bigdecimal::BigDecimal;
use chrono::NaiveDate;
use std::str::FromStr;

pub fn convert_date(field_value: Option<&&calamine::DataType>) -> Option<NaiveDate> {
    match field_value? {
        calamine::DataType::String(s) => {
            // if the fields are not separated by /, they are all together
            let v: Vec<&str> = s.split('/').collect();
            if v.len() == 3 {
                let day = u32::from_str(v.get(0)?).ok()?;
                let month = u32::from_str(v.get(1)?).ok()?;
                let year = i32::from_str(v.get(2)?).ok()?;
                Some(NaiveDate::from_ymd(year, month, day))
            } else if v.len() == 1 {
                let day = u32::from_str(s.get(0..2)?).ok()?;
                let month = u32::from_str(s.get(2..4)?).ok()?;
                let year = i32::from_str(s.get(4..8)?).ok()?;
                Some(NaiveDate::from_ymd(year, month, day))
            } else {
                None
            }
        }
        calamine::DataType::Float(f) => {
            // sometimes dates are encoded as floats in Excel
            // value 1899-12-30 + f days
            Some(NaiveDate::from_ymd(1899, 12, 30) + chrono::Duration::days(*f as i64))
        }
        _ => None,
    }
}

pub fn convert_string(field_value: Option<&&calamine::DataType>) -> Option<String> {
    match field_value {
        Some(calamine::DataType::String(s)) => Some(s.clone()),
        Some(calamine::DataType::Float(f)) => Some(format!("{}", f)),
        Some(calamine::DataType::Int(i)) => Some(format!("{}", i)),
        _ => None,
    }
}

pub fn convert_i32(field_value: Option<&&calamine::DataType>) -> Option<i32> {
    match field_value {
        Some(calamine::DataType::Float(f)) => Some(*f as i32),
        Some(calamine::DataType::Int(i)) => Some(*i as i32),
        _ => None,
    }
}

pub fn convert_decimal(field_value: Option<&&calamine::DataType>) -> Option<BigDecimal> {
    match field_value {
        Some(calamine::DataType::Float(f)) => Some(BigDecimal::from((*f * 100_f64) as i32) / 100),
        Some(calamine::DataType::Int(i)) => Some(BigDecimal::from(*i as i32)),
        Some(calamine::DataType::String(s)) => BigDecimal::from_str(s).ok(),
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
