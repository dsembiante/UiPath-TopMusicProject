# UiPath Top Music Project

A UiPath RPA bot that scrapes the top 10 songs from Amazon Music, Spotify, and YouTube, then automatically generates a formatted PowerPoint presentation for each platform — complete with song titles, artist names, album artwork, and lyrics for every track.

---

## What It Does

Most music chart comparisons require manually visiting multiple platforms, copying data, and formatting it yourself. This bot automates the entire process end to end — from scraping live chart data across three platforms to delivering a polished, slide-by-slide presentation for each one.

For each of the three platforms, the bot:
1. Navigates to the platform's top charts page
2. Scrapes the top 10 songs including title and artist name
3. Downloads the album artwork for each track
4. Extracts the lyrics for each song using JavaScript injection
5. Generates a PowerPoint presentation with one slide per song, containing the song title, artist, album art, and full lyrics

---

## Output

Three fully populated PowerPoint presentations, one per platform:

- `AmazonTopSongsPowerPoint.pptx`
- `SpotifyTopSongsPowerPoint.pptx`
- `YoutubeTopSongsPowerPoint.pptx`

Each presentation contains 10 slides. Every slide includes the song ranking, title, artist name, album artwork, and lyrics — all populated automatically without any manual input.

---

## Tech Stack

| Component | Technology | Purpose |
|---|---|---|
| Automation Platform | UiPath Studio (Community Edition) | Workflow orchestration and browser automation |
| Web Scraping | UiPath Browser Automation | Navigates and extracts data from Amazon, Spotify, and YouTube |
| JavaScript Injection | UiPath Inject JS Activity | Handles dynamic page elements that standard selectors cannot reach |
| Presentation Generation | UiPath PowerPoint Activities | Creates and populates slides with scraped content |
| Image Handling | UiPath Download File Activity | Downloads and stores album artwork locally |

---

## Project Structure

```
UiPath-TopMusicProject/
├── Main.xaml                        # Entry point — orchestrates all three platform workflows
├── ProjectSequencesFlowcharts/      # Sub-workflows for each platform scrape
├── AlbumArt/                        # Downloaded album artwork images
├── JS/                              # JavaScript injection scripts
├── Powerpoints/
│   ├── AmazonTopSongsPowerPoint.pptx
│   ├── SpotifyTopSongsPowerPoint.pptx
│   ├── YoutubeTopSongsPowerPoint.pptx
│   └── Template.pptx                # Base slide template
└── project.json                     # UiPath project configuration
```

---

## Key Technical Decisions

**Why JavaScript injection?**
Music platforms like Spotify and YouTube render content dynamically, meaning standard UiPath selectors cannot reliably target elements that load asynchronously. JavaScript injection bypasses this by executing directly against the DOM, making the scraper more resilient to page structure changes.

**Why separate workflows per platform?**
Each platform has a different page structure, selector pattern, and dynamic loading behavior. Isolating each into its own workflow in `ProjectSequencesFlowcharts/` keeps the logic clean, makes debugging straightforward, and allows any individual platform workflow to be updated independently without touching the others.

**Why PowerPoint instead of Excel?**
Album artwork and lyrics don't present well in a spreadsheet. PowerPoint allows each song to have its own dedicated slide with visual layout — making the output immediately presentable rather than requiring further formatting.

---

## How to Run

### Prerequisites
- UiPath Studio (Community Edition or higher)
- Google Chrome with UiPath Extension installed

### Steps
1. Clone the repo and open `project.json` in UiPath Studio
2. Ensure the Chrome UiPath extension is enabled
3. Run `Main.xaml`
4. Completed PowerPoint files will be saved to the `Powerpoints/` folder

---

## What I Would Add Next

- **Cross-platform comparison slide** — a summary slide at the end of each deck showing which songs appear on multiple charts simultaneously
- **Scheduled execution** — run the bot on a weekly schedule via UiPath Orchestrator to track chart movement over time
- **Email delivery** — automatically email the completed presentations to a distribution list after each run
- **Error handling improvements** — retry logic for dynamic pages that fail to load within the timeout window

---

## Author

Dominic Sembiante

LinkedIn: www.linkedin.com/in/dominic-sembiante-92b598149

GitHub: github.com/dsembiante

Built as part of an automation engineering portfolio demonstrating enterprise RPA development with UiPath, including multi-source web scraping, JavaScript injection, and automated document generation.
