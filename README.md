# 🎵 Random Thunder — Music League Archive

> *Every song. Every round. Every season. Searchable in seconds.*

---

## What Is This?

Random Thunder is a music discovery league where players submit songs each week based on a theme — and compete for glory, bragging rights, and the undying respect of their peers.

This is the **complete archive**: 4,595 submissions across 14 seasons, 253 rounds, and 43 players, compressed into a single searchable website that runs entirely in your browser.

---

## The Site

Open **`music_league.html`** in Chrome and start exploring.

### Features

- **Instant search** — type anything and results filter in real time. Song titles, artists, albums, round themes, submitter names — it all matches.
- **Season filters** — narrow to a specific season with one click.
- **Album art** — pulled live from iTunes for every card.
- **Spotify links** — click any card to open the song directly in Spotify.
- **Live stats** — songs, rounds, seasons, and players counted up on load.
- **Animated everything** — particle background, staggered card entrances, smooth transitions powered by [anime.js](https://animejs.com).

---

## Auto-Updates

A scheduled task runs **every Wednesday at 12:01 AM** — right after each new round drops. It:

1. Opens Music League in Chrome (uses your existing login session)
2. Scrapes the newest round's results
3. Adds any new songs to `Random_Thunder_Music_League.xlsx`
4. Regenerates `music_league.html` with the updated database

No manual work required. The archive stays current automatically.

> **First-time setup:** Open the Scheduled sidebar in Claude and click **Run now** on the `music-league-weekly-update` task. This pre-approves the browser permissions so future runs don't pause waiting for your input.

---

## Files

| File | Description |
|------|-------------|
| `music_league.html` | The searchable database site — open in Chrome |
| `Random_Thunder_Music_League.xlsx` | Source spreadsheet with all submissions |
| `update_music_league.py` | Script to merge new songs and regenerate the HTML |

---

*Built with Claude · Powered by anime.js · Data from Music League*

