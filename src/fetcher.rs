/// fetcher.rs
/// Downloads the raw ICS content from a USyd timetable subscription URL.

use anyhow::{Result, bail, Context};

/// Expected URL prefix for safety validation.
const EXPECTED_HOST: &str = "timetable.sydney.edu.au";

/// Download the ICS content from the given URL.
/// Validates that the URL points to the expected USyd timetable host.
pub fn fetch_ical(url: &str) -> Result<String> {
    // Basic URL validation
    validate_url(url)?;

    let response = ureq::get(url)
        .set("User-Agent", "iScheduler/0.1 (USyd timetable restructurer)")
        .call()
        .context("Failed to download the calendar. Please check the URL and your internet connection.")?;

    let status = response.status();
    if status != 200 {
        bail!(
            "Server returned HTTP {}. \
             Please verify your personal key in the URL is correct and hasn't expired.",
            status
        );
    }

    let content_type = response
        .header("content-type")
        .unwrap_or("")
        .to_lowercase();

    // Accept both text/calendar and text/plain (some servers send plain)
    if !content_type.contains("calendar")
        && !content_type.contains("text")
        && !content_type.is_empty()
    {
        eprintln!(
            "Warning: unexpected Content-Type '{}', attempting to parse anyway.",
            content_type
        );
    }

    let body = response
        .into_string()
        .context("Failed to read response body as UTF-8 text.")?;

    if body.is_empty() {
        bail!("Downloaded calendar is empty. The timetable may have no events enrolled.");
    }

    if !body.contains("BEGIN:VCALENDAR") {
        bail!(
            "Downloaded content does not look like a valid iCalendar file. \
             Please check the URL."
        );
    }

    Ok(body)
}

fn validate_url(url: &str) -> Result<()> {
    if url.trim().is_empty() {
        bail!("URL cannot be empty.");
    }

    if !url.starts_with("https://") && !url.starts_with("http://") {
        bail!("URL must start with https:// or http://");
    }

    if !url.contains(EXPECTED_HOST) {
        bail!(
            "URL does not appear to be a USyd timetable link.\n\
             Expected format: https://{}/even/rest/calendar/ical/<your-personal-key>",
            EXPECTED_HOST
        );
    }

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_validate_url_ok() {
        let url = "https://timetable.sydney.edu.au/even/rest/calendar/ical/abc123";
        assert!(validate_url(url).is_ok());
    }

    #[test]
    fn test_validate_url_wrong_host() {
        let url = "https://evil.example.com/ical";
        assert!(validate_url(url).is_err());
    }

    #[test]
    fn test_validate_url_no_scheme() {
        let url = "timetable.sydney.edu.au/even/rest/calendar/ical/abc";
        assert!(validate_url(url).is_err());
    }
}
