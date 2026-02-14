//! Excel serial date conversion.
//!
//! Excel stores dates as floating-point serial numbers: the integer part
//! counts days since 1899-12-30 (with the Lotus 1-2-3 bug that treats
//! 1900 as a leap year), and the fractional part is the time of day.
//!
//! Whether a numeric cell is a date depends on its number format (`numFmtId`).
//! This module provides both the format detection logic and the serial→ISO
//! conversion, shared by the `.xlsx` and `.xls` parsers.

/// Built-in numFmtIds that Excel defines as date/time formats.
///
/// Source: ECMA-376 Part 1, §18.8.30 (numFmt) and Microsoft documentation.
/// These IDs are hardcoded into Excel and never appear in styles.xml.
const BUILTIN_DATE_FMT_IDS: &[u16] = &[
    14, 15, 16, 17, 18, 19, 20, 21, 22, // standard date/time
    27, 28, 29, 30, 31, 32, 33, 34, 35, 36, // CJK date formats
    45, 46, 47, // time formats (mm:ss, [h]:mm:ss, mm:ss.0)
    50, 51, 52, 53, 54, 55, 56, 57, 58, // CJK extended date formats
];

/// Check if a `numFmtId` refers to a date/time format.
///
/// For built-in IDs, checks against Excel's hardcoded list.
/// For custom formats (ID ≥ 164), inspects the format string.
pub(crate) fn is_date_format_id(id: u16) -> bool {
    BUILTIN_DATE_FMT_IDS.contains(&id)
}

/// Check if a custom format string looks like a date/time format.
///
/// Heuristic: if the format contains date/time tokens (`y`, `m`, `d`, `h`, `s`)
/// but not number tokens (`0`, `#`, `?`), it's a date format. Ignores content
/// inside quoted strings and backslash-escaped characters.
pub(crate) fn is_date_format_string(fmt: &str) -> bool {
    let mut has_date_token = false;
    let mut has_number_token = false;
    let mut in_quote = false;
    let mut prev_backslash = false;

    for ch in fmt.chars() {
        if prev_backslash {
            prev_backslash = false;
            continue;
        }
        if ch == '\\' {
            prev_backslash = true;
            continue;
        }
        if ch == '"' {
            in_quote = !in_quote;
            continue;
        }
        if in_quote {
            continue;
        }

        match ch.to_ascii_lowercase() {
            // 'm' is ambiguous (month or minute) but in date context it's always date
            'y' | 'd' | 'h' | 's' | 'm' => has_date_token = true,
            '0' | '#' | '?' => has_number_token = true,
            _ => {}
        }
    }

    has_date_token && !has_number_token
}

/// Resolve which XF/style entries are date formats.
///
/// Given a list of `numFmtId` values (one per XF or cellXf entry) and a list
/// of custom format definitions `(numFmtId, format_code)`, returns a `Vec<bool>`
/// where each entry indicates whether that style index is a date format.
///
/// This logic is shared between the `.xlsx` (cellXfs) and `.xls` (XF records)
/// parsers, which both need to map style indices to date-or-not.
pub(crate) fn resolve_date_styles(fmt_ids: &[u16], custom_formats: &[(u16, String)]) -> Vec<bool> {
    fmt_ids
        .iter()
        .map(|&fmt_id| {
            if is_date_format_id(fmt_id) {
                return true;
            }
            custom_formats
                .iter()
                .any(|(id, code)| *id == fmt_id && is_date_format_string(code))
        })
        .collect()
}

/// Convert an Excel serial number to an ISO 8601 string.
///
/// Returns `YYYY-MM-DD` for whole numbers, `YYYY-MM-DD HH:MM:SS` if
/// there is a fractional (time) component.
///
/// Handles the Lotus 1-2-3 bug: serial 60 is treated as 1900-02-29
/// (which doesn't exist), and serials ≤ 0 or absurdly large values
/// are returned as-is.
pub(crate) fn serial_to_iso(serial: f64) -> String {
    if serial < 0.0 || serial.is_nan() || serial.is_infinite() {
        return format_fallback(serial);
    }

    #[allow(clippy::cast_possible_truncation)] // capped at 2_958_465 below
    let day_serial = serial.floor() as i64;
    let frac = serial - serial.floor();

    // Serial 0 is sometimes used as "no date"
    if day_serial == 0 {
        // Pure time value
        if frac > 0.0 {
            return format_time_only(frac);
        }
        return format_fallback(serial);
    }

    // Cap at year 9999 (~2_958_465)
    if day_serial > 2_958_465 {
        return format_fallback(serial);
    }

    let (year, month, day) = serial_to_ymd(day_serial);

    if frac.abs() < 1e-10 {
        format!("{year:04}-{month:02}-{day:02}")
    } else {
        let (hour, min, sec) = frac_to_hms(frac);
        format!("{year:04}-{month:02}-{day:02} {hour:02}:{min:02}:{sec:02}")
    }
}

/// Convert the integer part of a serial to (year, month, day).
///
/// Accounts for the Lotus 1-2-3 leap year bug: serial 60 is treated
/// as 1900-02-29. Serials > 60 are adjusted by -1 to compensate.
fn serial_to_ymd(serial: i64) -> (i32, u32, u32) {
    // Handle the fake 1900-02-29
    if serial == 60 {
        return (1900, 2, 29);
    }

    // Adjust for the Lotus bug: serials after 60 are off by one
    let adjusted = if serial > 60 { serial - 1 } else { serial };

    // serial 1 = 1900-01-01. Convert to days since a known epoch.
    // We'll convert to a Unix-like day count: 1970-01-01 = serial 25569.
    // Instead, just directly compute using cumulative day counts.

    // Days since 1900-01-01 (0-based: serial 1 → day 0)
    #[allow(clippy::cast_possible_truncation)] // max serial ~3M, fits in i32
    let mut days_remaining = (adjusted - 1) as i32;
    let mut year: i32 = 1900;

    // Advance by years
    loop {
        let days_in_year = if is_leap_year(year) { 366 } else { 365 };
        if days_remaining < days_in_year {
            break;
        }
        days_remaining -= days_in_year;
        year += 1;
    }

    // Advance by months
    let month_days: [i32; 12] = if is_leap_year(year) {
        [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    } else {
        [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    };

    let mut month: u32 = 1;
    for &md in &month_days {
        if days_remaining < md {
            break;
        }
        days_remaining -= md;
        month += 1;
    }

    let day = days_remaining.cast_unsigned() + 1;
    (year, month, day)
}

/// Check if a year is a leap year in the Gregorian calendar.
const fn is_leap_year(year: i32) -> bool {
    (year % 4 == 0 && year % 100 != 0) || year % 400 == 0
}

/// Convert the fractional part of a serial to (hour, minute, second).
fn frac_to_hms(frac: f64) -> (u32, u32, u32) {
    #[allow(clippy::cast_possible_truncation, clippy::cast_sign_loss)] // frac ∈ [0, 1)
    let total_seconds = (frac * 86400.0).round() as u64;
    let h = (total_seconds / 3600) % 24;
    let m = (total_seconds % 3600) / 60;
    let s = total_seconds % 60;
    #[allow(clippy::cast_possible_truncation)]
    (h as u32, m as u32, s as u32)
}

/// Format a pure time value (serial between 0 and 1).
fn format_time_only(frac: f64) -> String {
    let (h, m, s) = frac_to_hms(frac);
    format!("{h:02}:{m:02}:{s:02}")
}

/// Fallback: format as a number when the serial is out of range.
fn format_fallback(val: f64) -> String {
    if val.fract() == 0.0 && val.abs() < 1e15 {
        #[allow(clippy::cast_possible_truncation)]
        return format!("{}", val as i64);
    }
    val.to_string()
}

#[cfg(test)]
mod tests {
    use super::*;

    // ── is_date_format_id ─────────────────────────────────────────

    #[test]
    fn builtin_date_id_14() {
        assert!(is_date_format_id(14));
    }

    #[test]
    fn builtin_date_id_22() {
        assert!(is_date_format_id(22));
    }

    #[test]
    fn builtin_time_id_45() {
        assert!(is_date_format_id(45));
    }

    #[test]
    fn builtin_cjk_date_id_27() {
        assert!(is_date_format_id(27));
    }

    #[test]
    fn general_format_not_date() {
        assert!(!is_date_format_id(0));
    }

    #[test]
    fn number_format_not_date() {
        assert!(!is_date_format_id(1));
    }

    #[test]
    fn custom_format_base_not_date() {
        assert!(!is_date_format_id(164));
    }

    // ── is_date_format_string ─────────────────────────────────────

    #[test]
    fn standard_date_format() {
        assert!(is_date_format_string("yyyy-mm-dd"));
    }

    #[test]
    fn date_time_format() {
        assert!(is_date_format_string("yyyy-mm-dd hh:mm:ss"));
    }

    #[test]
    fn short_date_format() {
        assert!(is_date_format_string("m/d/yy"));
    }

    #[test]
    fn time_only_format() {
        assert!(is_date_format_string("hh:mm:ss"));
    }

    #[test]
    fn number_format_not_date_str() {
        assert!(!is_date_format_string("#,##0.00"));
    }

    #[test]
    fn general_format_not_date_str() {
        assert!(!is_date_format_string("General"));
    }

    #[test]
    fn percentage_not_date() {
        assert!(!is_date_format_string("0%"));
    }

    #[test]
    fn mixed_date_number_not_date() {
        // Contains both date tokens and number tokens — treat as number
        assert!(!is_date_format_string("yyyy-mm-dd #0"));
    }

    #[test]
    fn quoted_text_ignored() {
        // "d" inside quotes should not trigger date detection
        assert!(!is_date_format_string("\"day\""));
    }

    #[test]
    fn escaped_char_ignored() {
        assert!(!is_date_format_string("\\d"));
    }

    #[test]
    fn date_with_quoted_text() {
        assert!(is_date_format_string("yyyy\"年\"mm\"月\"dd\"日\""));
    }

    // ── serial_to_iso ─────────────────────────────────────────────

    #[test]
    fn epoch_day_one() {
        // Serial 1 = 1900-01-01
        assert_eq!(serial_to_iso(1.0), "1900-01-01");
    }

    #[test]
    fn lotus_bug_day_60() {
        // Serial 60 = the fake 1900-02-29
        assert_eq!(serial_to_iso(60.0), "1900-02-29");
    }

    #[test]
    fn day_after_lotus_bug() {
        // Serial 61 = 1900-03-01
        assert_eq!(serial_to_iso(61.0), "1900-03-01");
    }

    #[test]
    fn known_date_2024_01_01() {
        // 2024-01-01 = serial 45292
        assert_eq!(serial_to_iso(45292.0), "2024-01-01");
    }

    #[test]
    fn known_date_2000_01_01() {
        // 2000-01-01 = serial 36526
        assert_eq!(serial_to_iso(36526.0), "2000-01-01");
    }

    #[test]
    fn known_date_1999_12_31() {
        // 1999-12-31 = serial 36525
        assert_eq!(serial_to_iso(36525.0), "1999-12-31");
    }

    #[test]
    fn date_with_time() {
        // 2024-01-01 12:00:00 = 45292.5
        assert_eq!(serial_to_iso(45292.5), "2024-01-01 12:00:00");
    }

    #[test]
    fn date_with_time_6am() {
        // 2024-01-01 06:00:00 = 45292.25
        assert_eq!(serial_to_iso(45292.25), "2024-01-01 06:00:00");
    }

    #[test]
    fn pure_time_value() {
        // 0.5 = 12:00:00
        assert_eq!(serial_to_iso(0.5), "12:00:00");
    }

    #[test]
    fn pure_time_quarter_day() {
        // 0.25 = 06:00:00
        assert_eq!(serial_to_iso(0.25), "06:00:00");
    }

    #[test]
    fn negative_serial_fallback() {
        assert_eq!(serial_to_iso(-1.0), "-1");
    }

    #[test]
    fn zero_serial_fallback() {
        assert_eq!(serial_to_iso(0.0), "0");
    }

    #[test]
    fn huge_serial_fallback() {
        assert_eq!(serial_to_iso(3_000_000.0), "3000000");
    }

    #[test]
    fn known_date_feb_28_1900() {
        // Serial 59 = 1900-02-28
        assert_eq!(serial_to_iso(59.0), "1900-02-28");
    }

    #[test]
    fn known_date_mar_1_1900() {
        // Serial 61 = 1900-03-01 (because 60 is the fake Feb 29)
        assert_eq!(serial_to_iso(61.0), "1900-03-01");
    }

    #[test]
    fn known_date_dec_31_9999() {
        // Serial 2958465 = 9999-12-31 (max supported)
        assert_eq!(serial_to_iso(2_958_465.0), "9999-12-31");
    }

    // ── frac_to_hms ───────────────────────────────────────────────

    #[test]
    fn midnight() {
        assert_eq!(frac_to_hms(0.0), (0, 0, 0));
    }

    #[test]
    fn noon() {
        assert_eq!(frac_to_hms(0.5), (12, 0, 0));
    }

    #[test]
    fn end_of_day() {
        // Just under 1.0 — should round to 23:59:59 or 00:00:00
        assert_eq!(frac_to_hms(0.999_988_425_925_926), (23, 59, 59));
    }
}
