use std::num::ParseIntError;

fn get_column_letter(column_number: i32) -> String {
    let mut letter = String::new();
    let mut n = column_number;

    while n > 0 {
        let remainder = (n - 1) % 26;
        letter.insert(0, (remainder as u8 + b'A') as char);
        n = (n - remainder - 1) / 26;
    }

    letter
}

fn get_column_number(column_letter: &str) -> Result<i32, ParseIntError> {
    let mut col_num = 0;
    for (i, c) in column_letter.chars().rev().enumerate() {
        let digit = (c as u8 - b'A' + 1) as i32;
        col_num += digit * i32::pow(26, i as u32);
    }
    Ok(col_num)
}

pub fn get_cell_address(row_num: i32, col_num: i32) -> String {
    if row_num < 1 || col_num < 1 {
        panic!("Row and column must be positive integers");
    }

    let col_letter = get_column_letter(col_num);
    format!("{}{}", col_letter, row_num)
}

pub fn get_addr_int(cell_addr: &str) -> Result<(i32, i32), ParseIntError> {
    let re = regex::Regex::new(r"^([A-Z]+)(\d+)$").unwrap();
    let captures = re.captures(cell_addr).unwrap();
    let col_str = &captures[1];
    let row_str = &captures[2];
    let col_num = get_column_number(col_str)?;
    let row_num = row_str.parse::<i32>()?;
    Ok((row_num, col_num))
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::error::Error;
    #[test]
    fn test_get_addr_int() -> Result<(), Box<dyn Error>> {
        let cell_addr = String::from("AB123");
        let expected_row_num: i32 = 123;
        let expected_col_num: i32 = 28;

        let (row_num, col_num) = get_addr_int(&cell_addr)?;

        assert_eq!(row_num, expected_row_num);
        assert_eq!(col_num, expected_col_num);
        Ok(())
    }

    #[test]
    fn test_get_column_() -> Result<(), Box<dyn Error>> {
        let column_letters = [
            "A", "B", "C", "X", "Y", "Z", "AA", "AB", "AC", "ZY", "ZZ", "AAA", "AAB", "AAC", "ABC",
            "ABA",
        ];

        let column_numbers = [
            1, 2, 3, 24, 25, 26, 27, 28, 29, 701, 702, 703, 704, 705, 731, 729,
        ];

        for i in 0..column_letters.len() {
            let column_letter = column_letters[i];
            let expected_column_number = column_numbers[i];
            let actual_column_number = get_column_number(column_letter)?;
            assert_eq!(actual_column_number, expected_column_number);

            let expected_column_letter = column_letter.to_string();
            let actual_column_letter = get_column_letter(expected_column_number);
            assert_eq!(actual_column_letter, expected_column_letter);
        }

        Ok(())
    }

    #[test]
    fn test_get_cell_address() {
        assert_eq!(get_cell_address(1, 1), "A1");
        assert_eq!(get_cell_address(3, 26), "Z3");
        assert_eq!(get_cell_address(27, 27), "AA27");
        assert_eq!(get_cell_address(703, 137), "EG703");
    }
}
