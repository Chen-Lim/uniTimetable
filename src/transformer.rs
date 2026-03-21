/// transformer.rs
/// Transforms raw parsed events into clean, restructured `OutputEvent`s.
///
/// Key responsibilities:
/// 1. Parse and clean SUMMARY → "SubjectCode, ActivityType"
/// 2. Parse LOCATION dot-notation → human-readable building + room
/// 3. Group identical weekly events into RRULE recurrences
/// 4. Produce single-occurrence events for one-offs

use std::collections::HashMap;
use anyhow::{Result, Context};
use chrono::{NaiveDateTime, Duration, Datelike, Weekday};
use chrono_tz::Australia::Sydney;
use regex::Regex;
use uuid::Uuid;

use crate::parser::RawEvent;

// ── Data types ────────────────────────────────────────────────────────────────

/// A fully processed calendar event ready to be written to ICS.
#[derive(Debug, Clone)]
pub struct OutputEvent {
    pub uid: String,
    pub summary: String,
    pub description: String,
    pub location: Option<String>,
    pub dtstart: String, // formatted for ICS: YYYYMMDDTHHmmSS
    pub dtend: String,
    /// If Some, emit an RRULE repeating weekly until this date.
    pub rrule_until: Option<String>,
    /// Day abbreviation for RRULE BYDAY (MO, TU, WE, TH, FR, SA, SU)
    pub byday: Option<String>,
}

// ── Location parsing ──────────────────────────────────────────────────────────

/// Parse the USyd dot-separated location format.
///
/// Format: `{building_code}.{floor}.{room_id}.{Building Full Name}.{Room Full Name}`
/// Example: `H70.01.1040.Belinda Hutchinson Building.BHB Lecture Theatre 1040`
///
/// Returns `(building_address, room_name)` where building_address is
/// "{Building Full Name}, Sydney".
fn parse_location(raw: &str) -> Option<(String, String)> {
    let raw = raw.trim();
    if raw.is_empty() || raw == "-" {
        return None;
    }

    let parts: Vec<&str> = raw.splitn(6, '.').collect();
    //  idx:  0        1      2        3                    4
    //        H70    . 01  . 1040 . Belinda Hutchinson . BHB Lecture Theatre 1040

    let building = parts.get(3).map(|s| s.trim())?;
    let room = parts.get(4).map(|s| s.trim()).unwrap_or("Online");

    if building.is_empty() {
        return None;
    }

    Some((format!("{}, Sydney", building), room.to_string()))
}

// ── Summary parsing ───────────────────────────────────────────────────────────

/// Extract subject code from DESCRIPTION first line.
/// e.g. "COMP9208-S1C-NE-CC, Lecture, 01" → "COMP9208"
fn extract_subject_code(description: &str) -> Option<String> {
    let first_line = description.lines().next()?;
    let re = Regex::new(r"^([A-Z]{2,}\d{4,})").ok()?;
    re.captures(first_line)?.get(1).map(|m| m.as_str().to_string())
}

/// Extract activity type from DESCRIPTION first line.
/// e.g. "COMP9208-S1C-NE-CC, Lecture, 01" → "Lecture"
fn extract_activity_type(description: &str) -> Option<String> {
    let first_line = description.lines().next()?;
    // Format: "SUBJECTCODE-..., ActivityType, NN"
    let parts: Vec<&str> = first_line.splitn(3, ',').collect();
    parts.get(1).map(|s| s.trim().to_string())
}

// ── Datetime helpers ──────────────────────────────────────────────────────────

/// Parse an ICS datetime string (with or without TZID) to NaiveDateTime.
/// We treat all times as already in Australia/Sydney per the source calendar.
fn parse_dt(dt_str: &str) -> Result<NaiveDateTime> {
    // Strip trailing Z if present (treat as local)
    let s = dt_str.trim_end_matches('Z');
    NaiveDateTime::parse_from_str(s, "%Y%m%dT%H%M%S")
        .with_context(|| format!("Cannot parse datetime: {}", dt_str))
}

fn format_dt(dt: &NaiveDateTime) -> String {
    dt.format("%Y%m%dT%H%M%S").to_string()
}

fn weekday_abbr(wd: Weekday) -> &'static str {
    match wd {
        Weekday::Mon => "MO",
        Weekday::Tue => "TU",
        Weekday::Wed => "WE",
        Weekday::Thu => "TH",
        Weekday::Fri => "FR",
        Weekday::Sat => "SA",
        Weekday::Sun => "SU",
    }
}

// ── Grouping key ─────────────────────────────────────────────────────────────

/// Key used to group events that should become a single RRULE entry.
/// Two events are "the same recurring class" if they share summary, time-of-day,
/// and duration.
#[derive(Debug, PartialEq, Eq, Hash, Clone)]
struct GroupKey {
    summary: String,
    time_of_day: String, // "HHmmSS"
    duration_mins: i64,
    description_sig: String, // first line of description (identifies subject+group)
}

// ── Main transform ────────────────────────────────────────────────────────────

pub fn transform(raw_events: Vec<RawEvent>) -> Result<Vec<OutputEvent>> {
    // Group events by (summary, time, duration, description_signature)
    // to collapse weekly occurrences into RRULE.
    let mut groups: HashMap<GroupKey, Vec<RawEvent>> = HashMap::new();

    for event in raw_events {
        let summary = event.get("SUMMARY").unwrap_or("").to_string();
        let dtstart = event.get("DTSTART").unwrap_or("").to_string();
        let dtend = event.get("DTEND").unwrap_or("").to_string();
        let description = event.get("DESCRIPTION").unwrap_or("").to_string();

        if dtstart.is_empty() {
            continue;
        }

        let start_dt = match parse_dt(&dtstart) {
            Ok(d) => d,
            Err(e) => {
                eprintln!("Skipping event with bad DTSTART '{}': {}", dtstart, e);
                continue;
            }
        };
        let end_dt = match parse_dt(&dtend) {
            Ok(d) => d,
            Err(_) => start_dt + Duration::hours(1),
        };

        let duration_mins = (end_dt - start_dt).num_minutes();
        let time_of_day = start_dt.format("%H%M%S").to_string();
        let description_sig = description
            .lines()
            .next()
            .unwrap_or("")
            .to_string();

        let key = GroupKey {
            summary,
            time_of_day,
            duration_mins,
            description_sig,
        };

        groups.entry(key).or_default().push(event);
    }

    let mut output = Vec::new();

    for (key, mut occurrences) in groups {
        // Sort occurrences chronologically
        occurrences.sort_by_key(|e| e.get("DTSTART").unwrap_or("").to_string());

        // Use the first occurrence for shared metadata
        let first = &occurrences[0];
        let dtstart_str = first.get("DTSTART").unwrap_or("");
        let dtend_str   = first.get("DTEND").unwrap_or("");
        let description = first.get("DESCRIPTION").unwrap_or("").to_string();
        let raw_location = first.get("LOCATION").unwrap_or("-").to_string();

        let first_start = parse_dt(dtstart_str)?;
        let first_end = if dtend_str.is_empty() {
            first_start + Duration::minutes(key.duration_mins)
        } else {
            parse_dt(dtend_str).unwrap_or(first_start + Duration::minutes(key.duration_mins))
        };

        // Parse location once — shared across all segments
        let (location, room) = match parse_location(&raw_location) {
            Some((bldg, room)) => (Some(bldg), Some(room)),
            None => (None, None),
        };

        let subject_code = extract_subject_code(&description);
        let activity_type = extract_activity_type(&description);

        // Build summary: "COMP9208-Lecture" (matches correct ICS format)
        let clean_summary = match (&subject_code, &activity_type) {
            (Some(code), Some(activity)) => format!("{}-{}", code, activity),
            (Some(code), None) => code.clone(),
            _ => key.summary.clone(),
        };
        let clean_description = build_description(&description, room.as_deref());

        // ── Split into contiguous segments ────────────────────────────────────
        // If the gap between two consecutive occurrences is > 14 days (i.e. not
        // simply the normal weekly cadence), we start a new RRULE segment.
        // This correctly handles mid-semester breaks (Easter, study break, etc.)
        // where the timetable already exports separate VEVENT blocks.
        //
        // Each element of `segments` is a (start, end) pair of occurrence indices
        // into `occurrences`.
        let mut segments: Vec<(usize, usize)> = Vec::new();
        let mut seg_start = 0usize;

        for i in 1..occurrences.len() {
            let prev_dt = parse_dt(occurrences[i - 1].get("DTSTART").unwrap_or(""))?;
            let curr_dt = parse_dt(occurrences[i].get("DTSTART").unwrap_or(""))?;
            let gap = (curr_dt - prev_dt).num_days();

            if gap > 14 {
                // Gap larger than a single missed week → new segment
                segments.push((seg_start, i - 1));
                seg_start = i;
            }
        }
        segments.push((seg_start, occurrences.len() - 1));

        // ── Emit one OutputEvent per segment ──────────────────────────────────
        for (seg_idx, (from, to)) in segments.iter().enumerate() {
            let seg_first = &occurrences[*from];
            let seg_last  = &occurrences[*to];

            let seg_start_str = seg_first.get("DTSTART").unwrap_or("");
            let seg_end_str   = seg_first.get("DTEND").unwrap_or("");
            let seg_last_str  = seg_last.get("DTSTART").unwrap_or("");

            let seg_start_dt = parse_dt(seg_start_str)?;
            let seg_end_dt = if seg_end_str.is_empty() {
                seg_start_dt + Duration::minutes(key.duration_mins)
            } else {
                parse_dt(seg_end_str).unwrap_or(seg_start_dt + Duration::minutes(key.duration_mins))
            };
            let seg_until_dt = parse_dt(seg_last_str)?;

            let byday = weekday_abbr(seg_start_dt.weekday()).to_string();
            let count = to - from + 1; // number of occurrences in this segment

            // Only emit RRULE when there are multiple occurrences in the segment
            let (rrule_until, byday_field) = if count > 1 {
                (Some(format_dt(&seg_until_dt)), Some(byday))
            } else {
                (None, None)
            };

            // For the very first segment use the original description/location;
            // subsequent segments (after break) use the same metadata.
            let _ = seg_idx; // suppress unused warning — metadata is shared

            output.push(OutputEvent {
                uid: format!("{}@ischeduler.usyd", Uuid::new_v4()),
                summary: clean_summary.clone(),
                description: clean_description.clone(),
                location: location.clone(),
                dtstart: format_dt(&seg_start_dt),
                dtend: format_dt(&seg_end_dt),
                rrule_until,
                byday: byday_field,
            });
        }
    }

    // Sort output by DTSTART for a tidy file
    output.sort_by(|a, b| a.dtstart.cmp(&b.dtstart));

    Ok(output)
}

/// Build a clean, human-readable DESCRIPTION from the raw ICS DESCRIPTION value.
///
/// Raw example (after unescape):
///   COMP9208-S1C-NE-CC, Lecture, 01
///   Artificial Intelligence and Society
///   Staff: -
///   Location: H70.01.1040.Belinda Hutchinson Building.BHB Lecture Theatre 1040
///
/// We keep the course name and append the room if available.
fn build_description(raw: &str, room: Option<&str>) -> String {
    let mut lines: Vec<&str> = raw.lines().collect();

    // Remove the raw "Location:" line since we expose it in LOCATION field
    lines.retain(|l| !l.trim_start().starts_with("Location:"));

    let mut result = lines.join("\n");

    if let Some(r) = room {
        if !r.is_empty() && r != "Online" {
            result.push_str(&format!("\nRoom: {}", r));
        }
    }

    result
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_location_full() {
        let raw = "H70.01.1040.Belinda Hutchinson Building.BHB Lecture Theatre 1040";
        let (bldg, room) = parse_location(raw).unwrap();
        assert_eq!(bldg, "Belinda Hutchinson Building, Sydney");
        assert_eq!(room, "BHB Lecture Theatre 1040");
    }

    #[test]
    fn test_parse_location_dash() {
        assert!(parse_location("-").is_none());
        assert!(parse_location("").is_none());
    }

    #[test]
    fn test_extract_subject_code() {
        let desc = "COMP9208-S1C-NE-CC, Lecture, 01\nArtificial Intelligence";
        assert_eq!(extract_subject_code(desc), Some("COMP9208".to_string()));
    }

    #[test]
    fn test_extract_activity_type() {
        let desc = "COMP9208-S1C-NE-CC, Lecture, 01\nArtificial Intelligence";
        assert_eq!(extract_activity_type(desc), Some("Lecture".to_string()));
        let desc2 = "PMGT5850-S1C-NE-CC, Tutorial, 01\nCapstone";
        assert_eq!(extract_activity_type(desc2), Some("Tutorial".to_string()));
    }

    #[test]
    fn test_weekday_abbr() {
        assert_eq!(weekday_abbr(Weekday::Mon), "MO");
        assert_eq!(weekday_abbr(Weekday::Fri), "FR");
    }
}
