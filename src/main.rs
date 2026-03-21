/// main.rs
/// iScheduler – USyd timetable iCal restructurer
///
/// Usage:
///   ischeduler --url "https://timetable.sydney.edu.au/even/rest/calendar/ical/<key>"
///   ischeduler --url "..." --output my_timetable.ics
///   ischeduler --file raw_timetable.ics           # process a local file instead

mod fetcher;
mod parser;
mod transformer;
mod writer;

use std::path::PathBuf;
use anyhow::{Result, bail};
use clap::Parser;

/// iScheduler: restructures your USyd iCal subscription into a clean calendar file.
#[derive(Parser, Debug)]
#[command(
    name = "ischeduler",
    version = "0.1.0",
    about = "Restructures USyd timetable iCal feeds into clean, import-ready ICS files.",
    long_about = None
)]
struct Cli {
    /// Your USyd iCal subscription URL.
    /// Format: https://timetable.sydney.edu.au/even/rest/calendar/ical/<your-personal-key>
    #[arg(short, long, group = "input")]
    url: Option<String>,

    /// Path to a locally downloaded ICS file (alternative to --url).
    #[arg(short, long, group = "input")]
    file: Option<PathBuf>,

    /// Output file path (default: timetable_clean.ics in the current directory).
    #[arg(short, long, default_value = "timetable_clean.ics")]
    output: PathBuf,

    /// Print verbose processing information.
    #[arg(short, long)]
    verbose: bool,
}

fn main() {
    if let Err(e) = run() {
        eprintln!("\n❌ Error: {:#}", e);
        std::process::exit(1);
    }
}

fn run() -> Result<()> {
    let cli = Cli::parse();

    // ── Step 1: Obtain raw ICS content ────────────────────────────────────────
    let raw_ics = if let Some(url) = &cli.url {
        println!("📡 Downloading timetable from USyd...");
        let content = fetcher::fetch_ical(url)?;
        if cli.verbose {
            println!("   Downloaded {} bytes.", content.len());
        }
        content
    } else if let Some(path) = &cli.file {
        println!("📂 Reading local file: {}", path.display());
        std::fs::read_to_string(path)
            .map_err(|e| anyhow::anyhow!("Cannot read file '{}': {}", path.display(), e))?
    } else {
        bail!(
            "Please provide either --url or --file.\n\
             Run with --help for usage information."
        );
    };

    // ── Step 2: Parse raw ICS ─────────────────────────────────────────────────
    println!("🔍 Parsing calendar events...");
    let raw_events = parser::parse_ics(&raw_ics)?;

    if cli.verbose {
        println!("   Found {} raw VEVENT blocks.", raw_events.len());
    }

    // ── Step 3: Transform ─────────────────────────────────────────────────────
    println!("⚙️  Restructuring events...");
    let output_events = transformer::transform(raw_events)?;

    if cli.verbose {
        println!(
            "   Produced {} output events ({} with RRULE recurrence).",
            output_events.len(),
            output_events.iter().filter(|e| e.rrule_until.is_some()).count()
        );
        println!("\n   Event list:");
        for ev in &output_events {
            let recur = if ev.rrule_until.is_some() { "↻" } else { "·" };
            println!("   {} {} ({})", recur, ev.summary, ev.dtstart);
        }
    }

    // ── Step 4: Write output ICS ──────────────────────────────────────────────
    println!("💾 Writing output to: {}", cli.output.display());
    writer::write_ics(&output_events, &cli.output)?;

    println!(
        "\n✅ Done! {} events written to '{}'.",
        output_events.len(),
        cli.output.display()
    );
    println!(
        "   Import this file into Apple Calendar, Google Calendar, or Outlook."
    );

    Ok(())
}
