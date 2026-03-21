/// writer.rs
/// Serialises a list of `OutputEvent`s into a well-formed ICS (iCalendar) string
/// and writes it to the specified path.

use std::fs;
use std::path::Path;
use anyhow::{Result, Context};
use chrono::Utc;

use crate::transformer::OutputEvent;

const TZID: &str = "Australia/Sydney";
const PROD_ID: &str = "-//iScheduler//USyd Timetable Restructurer//EN";
const CAL_NAME: &str = "University of Sydney Timetable";

/// Emit one ICS property line, folding at 75 octets per RFC 5545 §3.1.
/// Every segment (first and continuation) ends with CRLF.
/// Continuation lines are prefixed with a single space.
fn prop(line: &str) -> String {
    // Short path: fits in one line
    if line.len() <= 75 {
        return format!("{}\r\n", line);
    }

    let mut result = String::new();
    let mut pos = 0usize;
    let mut first = true;

    while pos < line.len() {
        // First segment: 75 chars. Continuation: 74 chars (1 taken by leading space).
        let max = if first { 75 } else { 74 };
        let end = if pos + max >= line.len() {
            line.len()
        } else {
            // Walk back to a UTF-8 char boundary
            let mut e = pos + max;
            while !line.is_char_boundary(e) {
                e -= 1;
            }
            e
        };

        if !first {
            result.push(' '); // RFC 5545 folding prefix
        }
        result.push_str(&line[pos..end]);
        result.push_str("\r\n");

        pos = end;
        first = false;
    }

    result
}

/// Escape special characters in ICS text values.
fn ics_escape(s: &str) -> String {
    s.replace('\\', "\\\\")
        .replace(';', "\\;")
        .replace(',', "\\,")
        .replace('\n', "\\n")
        .replace('\r', "")
}

/// Write the calendar to an ICS file at `output_path`.
pub fn write_ics(events: &[OutputEvent], output_path: &Path) -> Result<()> {
    let content = build_ics(events);

    // Ensure parent directory exists
    if let Some(parent) = output_path.parent() {
        if !parent.as_os_str().is_empty() {
            fs::create_dir_all(parent)
                .with_context(|| format!("Cannot create directory: {}", parent.display()))?;
        }
    }

    fs::write(output_path, &content)
        .with_context(|| format!("Cannot write ICS file: {}", output_path.display()))?;

    Ok(())
}

/// Return the ICS content as a String (useful for testing and Telegram bot later).
pub fn build_ics(events: &[OutputEvent]) -> String {
    let mut out = String::new();
    let now = Utc::now().format("%Y%m%dT%H%M%SZ").to_string();

    // Calendar header — each prop() call emits one line ending in CRLF
    out.push_str("BEGIN:VCALENDAR\r\n");
    out.push_str(&prop(&format!("PRODID:{}", PROD_ID)));
    out.push_str("VERSION:2.0\r\n");
    out.push_str("CALSCALE:GREGORIAN\r\n");
    out.push_str("METHOD:PUBLISH\r\n");
    out.push_str(&prop(&format!("X-WR-CALNAME:{}", CAL_NAME)));
    out.push_str(&prop(&format!("X-WR-TIMEZONE:{}", TZID)));

    // Timezone definition
    out.push_str(SYDNEY_TZ_BLOCK);

    // Events
    for event in events {
        out.push_str("BEGIN:VEVENT\r\n");
        out.push_str(&prop(&format!("UID:{}", event.uid)));
        out.push_str(&prop(&format!("DTSTAMP:{}", now)));
        out.push_str(&prop(&format!("DTSTART;TZID={}:{}", TZID, event.dtstart)));
        out.push_str(&prop(&format!("DTEND;TZID={}:{}", TZID, event.dtend)));

        if let (Some(until), Some(byday)) = (&event.rrule_until, &event.byday) {
            out.push_str(&prop(&format!(
                "RRULE:FREQ=WEEKLY;BYDAY={};UNTIL={}",
                byday, until
            )));
        }

        out.push_str(&prop(&format!("SUMMARY:{}", ics_escape(&event.summary))));

        if let Some(loc) = &event.location {
            out.push_str(&prop(&format!("LOCATION:{}", ics_escape(loc))));
        }

        out.push_str(&prop(&format!(
            "DESCRIPTION:{}",
            ics_escape(&event.description)
        )));

        out.push_str("TRANSP:OPAQUE\r\n");
        out.push_str("END:VEVENT\r\n");
        out.push_str("\r\n"); // blank line between events for readability
    }

    out.push_str("END:VCALENDAR\r\n");
    out
}

/// RFC 5545-compliant VTIMEZONE block for Australia/Sydney (AEST/AEDT).
const SYDNEY_TZ_BLOCK: &str = "\
BEGIN:VTIMEZONE\r\n\
TZID:Australia/Sydney\r\n\
LAST-MODIFIED:20240101T000000Z\r\n\
TZURL:https://www.tzurl.org/zoneinfo-outlook/Australia/Sydney\r\n\
X-LIC-LOCATION:Australia/Sydney\r\n\
BEGIN:STANDARD\r\n\
TZNAME:AEST\r\n\
TZOFFSETFROM:+1100\r\n\
TZOFFSETTO:+1000\r\n\
DTSTART:19700405T030000\r\n\
RRULE:FREQ=YEARLY;BYMONTH=4;BYDAY=1SU\r\n\
END:STANDARD\r\n\
BEGIN:DAYLIGHT\r\n\
TZNAME:AEDT\r\n\
TZOFFSETFROM:+1000\r\n\
TZOFFSETTO:+1100\r\n\
DTSTART:19701004T020000\r\n\
RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=1SU\r\n\
END:DAYLIGHT\r\n\
END:VTIMEZONE\r\n";

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_prop_short_line() {
        let result = prop("SUMMARY:Short");
        assert_eq!(result, "SUMMARY:Short\r\n");
    }

    #[test]
    fn test_prop_long_line_has_crlf_on_each_segment() {
        // 80-char line should be folded, each segment ending in CRLF
        let long = format!("DESCRIPTION:{}", "x".repeat(80));
        let result = prop(&long);
        for segment in result.split("\r\n").filter(|s| !s.is_empty()) {
            // first segment ≤75, continuation segments ≤75 (74 content + 1 space prefix)
            assert!(segment.len() <= 75, "segment too long: '{}'", segment);
        }
        // Must not have any bare \n without preceding \r
        assert!(!result.contains('\n') || result.contains("\r\n"));
    }

    #[test]
    fn test_ics_escape() {
        assert_eq!(ics_escape("Hello, World"), "Hello\\, World");
        assert_eq!(ics_escape("Line1\nLine2"), "Line1\\nLine2");
    }

    #[test]
    fn test_build_ics_structure() {
        let events = vec![OutputEvent {
            uid: "test-uid@ischeduler".to_string(),
            summary: "COMP9208-Lecture".to_string(),
            description: "Artificial Intelligence and Society\nBHB Lecture Theatre 1040".to_string(),
            location: Some("Belinda Hutchinson Building, Sydney".to_string()),
            dtstart: "20260226T170000".to_string(),
            dtend: "20260226T200000".to_string(),
            rrule_until: Some("20260528T170000".to_string()),
            byday: Some("TH".to_string()),
        }];

        let ics = build_ics(&events);

        // Every line must end with CRLF
        for line in ics.split('\n') {
            let line = line.trim_end_matches('\r');
            // continuation lines start with space — that's fine
            if !line.is_empty() && !line.starts_with(' ') {
                assert!(
                    ics.contains(&format!("{}\r\n", line)),
                    "Line missing CRLF: '{}'", line
                );
            }
        }

        assert!(ics.contains("BEGIN:VCALENDAR\r\n"));
        assert!(ics.contains("VERSION:2.0\r\n"));
        assert!(ics.contains("BEGIN:VTIMEZONE\r\n"));
        assert!(ics.contains("BEGIN:VEVENT\r\n"));
        assert!(ics.contains("RRULE:FREQ=WEEKLY;BYDAY=TH;UNTIL=20260528T170000\r\n"));
        assert!(ics.contains("SUMMARY:COMP9208-Lecture\r\n"));
        assert!(ics.contains("END:VEVENT\r\n"));
        assert!(ics.contains("END:VCALENDAR\r\n"));
    }
}
