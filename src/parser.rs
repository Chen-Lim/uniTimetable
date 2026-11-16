//! parser.rs
//! Parses a raw ICS string into a list of structured `RawEvent` maps.
//! Handles RFC 5545 line folding (lines starting with whitespace are continuations).

use std::collections::HashMap;
use anyhow::{Result, bail};

/// A raw calendar event: property name → value (unescaped).
/// Multi-valued properties (like multiple DESCRIPTION lines) use the last value,
/// which matches typical USyd timetable output.
#[derive(Debug, Clone)]
pub struct RawEvent {
    pub properties: HashMap<String, String>,
}

impl RawEvent {
    pub fn get(&self, key: &str) -> Option<&str> {
        self.properties.get(key).map(|s| s.as_str())
    }
}

/// Unfold RFC 5545 folded lines.
/// A folded line is a CRLF or LF followed by a single whitespace character.
fn unfold(input: &str) -> String {
    // Replace CRLF+whitespace or LF+whitespace continuations
    let crlf_folded = input.replace("\r\n ", "").replace("\r\n\t", "");
    crlf_folded.replace("\n ", "").replace("\n\t", "")
}

/// Unescape ICS text values (\n → newline, \, → ,, \; → ;, \\ → \) in a single pass.
fn unescape(value: &str) -> String {
    let mut out = String::with_capacity(value.len());
    let mut chars = value.chars();
    while let Some(c) = chars.next() {
        if c == '\\' {
            match chars.next() {
                Some('n') | Some('N') => out.push('\n'),
                Some(',') => out.push(','),
                Some(';') => out.push(';'),
                Some('\\') => out.push('\\'),
                Some(other) => {
                    out.push('\\');
                    out.push(other);
                }
                None => out.push('\\'),
            }
        } else {
            out.push(c);
        }
    }
    out
}

/// Split a property line into (name_with_params, value).
/// Handles lines like:
///   DTSTART;TZID=Australia/Sydney:20260226T170000
///   SUMMARY:Artificial Intelligence and Society\, Lecture
fn split_property(line: &str) -> Option<(String, String)> {
    // Find the first colon that separates name+params from value
    let colon_pos = line.find(':')?;
    let name_part = &line[..colon_pos];
    let value_part = &line[colon_pos + 1..];

    // Normalise parameter names: strip params (e.g. ";TZID=...") to get base name
    let base_name = if let Some(semi) = name_part.find(';') {
        name_part[..semi].to_uppercase()
    } else {
        name_part.to_uppercase()
    };

    Some((base_name, unescape(value_part)))
}

/// Parse the full ICS text and return all VEVENT blocks as `RawEvent`s.
pub fn parse_ics(raw: &str) -> Result<Vec<RawEvent>> {
    let unfolded = unfold(raw);
    let mut events = Vec::new();
    let mut in_event = false;
    let mut current: HashMap<String, String> = HashMap::new();

    for line in unfolded.lines() {
        let line = line.trim_end_matches('\r');

        match line {
            "BEGIN:VEVENT" => {
                in_event = true;
                current = HashMap::new();
            }
            "END:VEVENT" => {
                if in_event {
                    events.push(RawEvent { properties: std::mem::take(&mut current) });
                    in_event = false;
                }
            }
            _ if in_event => {
                if let Some((key, value)) = split_property(line) {
                    // For DESCRIPTION we want to keep the raw value intact
                    current.insert(key, value);
                }
            }
            _ => {}
        }
    }

    if events.is_empty() {
        bail!("No VEVENT blocks found in the downloaded calendar. \
               Please check that the URL is correct and the calendar is not empty.");
    }

    Ok(events)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_unfold() {
        let folded = "DESCRIPTION:Hello\r\n World";
        assert_eq!(unfold(folded), "DESCRIPTION:HelloWorld");
    }

    #[test]
    fn test_split_property_with_params() {
        let line = "DTSTART;TZID=Australia/Sydney:20260226T170000";
        let (name, value) = split_property(line).unwrap();
        assert_eq!(name, "DTSTART");
        assert_eq!(value, "20260226T170000");
    }

    #[test]
    fn test_unescape() {
        assert_eq!(unescape("Hello\\nWorld"), "Hello\nWorld");
        assert_eq!(unescape("A\\,B\\;C"), "A,B;C");
        assert_eq!(unescape("\\\\n"), "\\n");
        assert_eq!(unescape("A\\\\B"), "A\\B");
    }

    #[test]
    fn test_parse_ics_basic() {
        let raw = "\
BEGIN:VCALENDAR\r\n\
BEGIN:VEVENT\r\n\
DTSTART;TZID=Australia/Sydney:20260226T170000\r\n\
SUMMARY:COMP9208\\, Lecture\r\n\
DESCRIPTION:Line1\\n\r\n Line2\r\n\
END:VEVENT\r\n\
END:VCALENDAR\r\n";
        let events = parse_ics(raw).unwrap();
        assert_eq!(events.len(), 1);
        assert_eq!(events[0].get("SUMMARY"), Some("COMP9208, Lecture"));
        assert_eq!(events[0].get("DESCRIPTION"), Some("Line1\nLine2"));
    }
}
