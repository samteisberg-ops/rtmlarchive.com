#!/usr/bin/env python3
"""
Music League Weekly Updater
Run after scraping new round data to update the xlsx and HTML.

Usage:
  python3 update_music_league.py --new-songs '[ {"league":"...", "round":"...", ...}, ... ]'

Or call update_from_list(new_songs) from another script.
"""

import json, sys, os, re
from pathlib import Path

BASE = Path(__file__).parent
XLSX = BASE / "Random_Thunder_Music_League.xlsx"
VOTES_XLSX = BASE / "Random_Thunder_Music_League_Votes.xlsx"
HTML = BASE / "music_league.html"
VOTES_HTML = BASE / "music_league_votes.html"
ANALYTICS_JSON = BASE / "rtml_analytics.json"


# ── Safe block extraction/splicing ──────────────────────────────────────
#
# The old code used `re.sub(r'const X=\[.*?\];', ..., flags=re.DOTALL)` to
# splice fresh data into the HTML. That pattern is lazy ("first match
# wins") and has two failure modes that can silently corrupt the file:
#   1. If the marker doesn't match at all (e.g. spacing drifts between
#      "const X=[" and "const X = ["), re.sub just returns the string
#      UNCHANGED — no error, no warning; the script reports success with
#      stale data baked in.
#   2. The lazy ".*?\];" matches the FIRST "];" it finds. If that text
#      happens inside the data itself, the splice can land in the wrong
#      place, corrupting/truncating everything after it.
#
# _find_block() does a proper bracket-depth walk with string/escape
# awareness instead, so it can't be fooled by "]" or ";" inside string
# literals, and raises loudly (BlockNotFound) rather than silently
# no-op'ing if a marker can't be located.

class BlockNotFound(RuntimeError):
    pass


def _find_block(text, marker, open_char='[', close_char=']'):
    """Find `marker` then return (start, end) spanning the full
    "marker...[...];" or "marker...{...};" statement, end being just past
    the semicolon. Raises BlockNotFound instead of silently failing.
    Use open_char='{', close_char='}' for JS object literals."""
    start = text.find(marker)
    if start == -1:
        raise BlockNotFound(f"marker {marker!r} not found in document")
    bracket_start = text.find(open_char, start)
    if bracket_start == -1:
        raise BlockNotFound(f"no {open_char!r} found after marker {marker!r}")
    depth = 0
    in_str = None
    escape = False
    i = bracket_start
    n = len(text)
    while i < n:
        ch = text[i]
        if in_str:
            if escape:
                escape = False
            elif ch == '\\':
                escape = True
            elif ch == in_str:
                in_str = None
        else:
            if ch in ('"', "'", '`'):
                in_str = ch
            elif ch == open_char:
                depth += 1
            elif ch == close_char:
                depth -= 1
                if depth == 0:
                    end = i + 1
                    if end < n and text[end] == ';':
                        end += 1
                    return start, end
        i += 1
    raise BlockNotFound(f"unbalanced brackets after marker {marker!r} (reached end of file)")


def _splice_block(text, marker, new_statement, open_char='[', close_char=']'):
    """Replace the "marker...[...];" or "marker...{...};" statement with
    new_statement. Raises BlockNotFound (loudly) instead of the old silent
    no-op. Pass open_char='{', close_char='}' for JS object literals."""
    start, end = _find_block(text, marker, open_char, close_char)
    return text[:start] + new_statement + text[end:]


def validate_html(text, label=""):
    """Return a list of problems found in a regenerated HTML doc, or []
    if it looks structurally sound. Run before any regenerated file is
    written/pushed, to catch truncation or a botched splice up front
    instead of finding out when the live site breaks."""
    problems = []
    stripped = text.rstrip()
    if not stripped.endswith('</html>'):
        tail = stripped[-60:].replace('\n', '\\n')
        problems.append(f"{label}: file does not end with </html> (possible truncation); tail=...{tail!r}")
    open_script = len(re.findall(r'<script\b', text))
    close_script = text.count('</script>')
    if open_script != close_script:
        problems.append(f"{label}: <script> tag mismatch: {open_script} open vs {close_script} close")
    if '</body>' not in text:
        problems.append(f"{label}: missing </body>")
    for marker, name in [
        ("const SONGS=", "SONGS"),
        ("const PLAYERS = ", "PLAYERS"),
        ("const PLAYER_STATS = ", "PLAYER_STATS"),
        ("const ROUND_SUMMARIES = ", "ROUND_SUMMARIES"),
        ("const SONG_VOTES = ", "SONG_VOTES"),
    ]:
        if marker not in text:
            continue
        try:
            start, end = _find_block(text, marker)
        except BlockNotFound as e:
            problems.append(f"{label}: {name} block malformed: {e}")
            continue
        block = text[start:end]
        bracket_text = block[block.find('['):block.rfind(']') + 1]
        try:
            json.loads(bracket_text)
        except json.JSONDecodeError as e:
            problems.append(f"{label}: {name} block is not valid JSON: {e}")
    return problems


def _write_validated(path, text, label):
    """Write `text` to `path` only after validating it, via an atomic
    temp-file + os.replace so a crash mid-write never leaves a
    half-written file in place. Keeps a .bak of the previous good file.
    Raises RuntimeError if validation fails — the file on disk is left
    untouched in that case."""
    problems = validate_html(text, label)
    if problems:
        raise RuntimeError(
            f"Refusing to write {path.name} — validation failed:\n  " + "\n  ".join(problems)
        )
    if path.exists():
        backup = path.with_suffix(path.suffix + ".bak")
        backup.write_text(path.read_text(encoding="utf-8"), encoding="utf-8")
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(text, encoding="utf-8")
    os.replace(str(tmp), str(path))


def _compute_hof_data(songs):
    """Compute Hall of Fame data from the full songs list.

    Returns a dict with:
      hof_official  — dict: league → [[name, score], [name, score], [name, score]]
                      Top 3 players by total score within each season.
      in_progress   — the latest (highest-numbered) season name, marked as
                      still-accumulating in the UI.
      season_order  — leagues sorted chronologically (S1 first, current last).
      season_labels — dict: league → "Season N" display label.

    Only the `in_progress` season's entry is live-computed; all other
    seasons' entries are derived from the same xlsx scores that have always
    driven the HOF, so they stay stable and consistent with history.
    """
    from collections import defaultdict

    # Tally total score per (league, submitter) across all songs
    totals = defaultdict(lambda: defaultdict(int))
    for s in songs:
        league = s.get("league", "")
        submitter = s.get("submitter", "")
        score = s.get("score", 0) or 0
        if league and submitter:
            totals[league][submitter] += score

    # Determine chronological season order
    def _season_num(league):
        if league == "Random Thunder":
            return 1
        m = re.search(r'S(\d+)$', league)
        return int(m.group(1)) if m else 0

    season_order = sorted(totals.keys(), key=_season_num)
    in_progress = season_order[-1] if season_order else ""

    # Build SEASON_LABELS
    season_labels = {}
    for league in season_order:
        n = _season_num(league)
        season_labels[league] = f"Season {n}"

    # Build HOF_OFFICIAL: top 3 per season
    hof_official = {}
    for league in season_order:
        top3 = sorted(totals[league].items(), key=lambda x: -x[1])[:3]
        hof_official[league] = [[name, score] for name, score in top3]

    return {
        "hof_official": hof_official,
        "in_progress": in_progress,
        "season_order": season_order,
        "season_labels": season_labels,
        "totals": totals,
        "season_labels_all": season_labels,
    }


def _compute_season_finale(hof, league):
    """Build the SEASON_FINALE payload (winner + podium) for a season that
    just wrapped up, from the same totals _compute_hof_data already tallied.
    Returns None if the league has no data."""
    podium = hof["hof_official"].get(league)
    if not podium:
        return None
    return {
        "league": league,
        "seasonLabel": hof["season_labels"].get(league, league),
        "winner": podium[0][0],
        "score": podium[0][1],
        "podium": podium,
    }


def update_from_list(new_songs, analytics=None, season_complete=None):
    """
    new_songs: list of dicts with keys:
      league, round, roundNum, song, artist, album, score, place, submitter
    analytics (optional): dict with keys:
      players, playerStats, roundSummaries, songVotes
      NOTE: pass only NEW/changed data here, not the full existing+new arrays.
      _save_analytics() merges it into the existing JSON safely (append/update,
      never blind-replace). songVotes specifically should contain one vote
      array per entry in new_songs, in the same order, so it can be aligned
      to whichever of those songs actually get added (post-dedup).
    season_complete (optional): league name (e.g. "Random Thunder S14") whose
      final round was just added. When set, the regenerated HTML embeds a
      SEASON_FINALE payload (winner + podium) that triggers the "season
      complete" popup banner on next page load for anyone who hasn't
      dismissed it yet. Leave unset for a normal weekly update — the
      previously-set SEASON_FINALE (if any) is left untouched.
    """
    import openpyxl
    wb = openpyxl.load_workbook(str(XLSX))
    ws = wb.active

    # Build set of existing (league, round, song, artist) to avoid dupes
    existing = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        league, rnd, rnum, song, artist = row[0], row[1], row[2], row[3], row[4]
        existing.add((str(league), str(rnd), str(song), str(artist)))

    added = 0
    added_indices = []  # positions within new_songs that actually got added
    for idx, s in enumerate(new_songs):
        key = (str(s.get('league','')), str(s.get('round','')), str(s.get('song','')), str(s.get('artist','')))
        if key in existing:
            continue
        ws.append([
            s.get('league',''), s.get('round',''), s.get('roundNum',0),
            s.get('song',''), s.get('artist',''), s.get('album',''),
            s.get('score',0), s.get('place',''), s.get('submitter','')
        ])
        existing.add(key)
        added += 1
        added_indices.append(idx)

    wb.save(str(XLSX))
    print(f"Added {added} new songs to {XLSX}")

    if analytics:
        _save_analytics(analytics, added_indices)

    if added_indices:
        _update_votes_xlsx(new_songs, added_indices, analytics)

    if added > 0 or analytics or season_complete:
        regenerate_html(season_complete=season_complete)
    return added


def _update_votes_xlsx(new_songs, added_indices, analytics):
    """Append rows for newly-added songs to Random_Thunder_Music_League_Votes.xlsx,
    keeping it in sync with the main xlsx on every weekly run.

    Columns: League, Round, Round #, Song, Artist, Album, Official Score,
    Raw Upvotes, Raw Downvotes, Place, Submitter, then one column per player
    holding the exact point value they cast for that song (blank if they
    didn't vote). Raw Upvotes/Downvotes are COUNTS of positive/negative
    votes, not summed magnitude.

    analytics['songVotes'], if supplied, must be aligned to added_indices
    (one vote-array per song that was actually added, in order) using
    player indices into the merged players list — same convention as
    _save_analytics. If analytics/songVotes is missing, rows are still
    appended with the base columns filled in and vote columns left blank
    (best effort, matches the songs-only update path).

    New voter/submitter names not yet present as columns are added as new
    columns automatically (existing rows get a blank cell there).
    """
    if not VOTES_XLSX.exists():
        print(f"{VOTES_XLSX.name} not found, skipping votes xlsx update")
        return

    import openpyxl
    wb = openpyxl.load_workbook(str(VOTES_XLSX))
    ws = wb["Votes"] if "Votes" in wb.sheetnames else wb.active

    BASE_COLS = ["League", "Round", "Round #", "Song", "Artist", "Album",
                 "Official Score", "Raw Upvotes", "Raw Downvotes", "Place", "Submitter"]

    header = [c.value for c in ws[1]]
    if not header:
        header = list(BASE_COLS)
        for i, h in enumerate(header, start=1):
            ws.cell(row=1, column=i).value = h

    # Map added-song index -> list of (voterName, points), using the fresh
    # merged players list written to disk by _save_analytics (if it ran).
    song_votes_by_added = {}
    if analytics and analytics.get("songVotes"):
        sv = analytics["songVotes"]
        if len(sv) == len(added_indices):
            players_full = []
            if ANALYTICS_JSON.exists():
                players_full = json.loads(ANALYTICS_JSON.read_text(encoding="utf-8")).get("players", [])
            for pos, orig_idx in enumerate(added_indices):
                named = []
                for pidx, pts in sv[pos]:
                    if isinstance(pidx, int) and 0 <= pidx < len(players_full):
                        named.append((players_full[pidx], pts))
                song_votes_by_added[orig_idx] = named
        else:
            print(f"WARNING: songVotes length ({len(sv)}) != added songs ({len(added_indices)}) "
                  f"— votes xlsx rows will have blank vote columns")

    # Add any new voter/submitter columns that don't exist yet
    existing_cols = set(header[len(BASE_COLS):])
    needed_names = set()
    for orig_idx in added_indices:
        s = new_songs[orig_idx]
        if s.get("submitter"):
            needed_names.add(s["submitter"])
        for name, _pts in song_votes_by_added.get(orig_idx, []):
            needed_names.add(name)
    new_cols = sorted(n for n in needed_names if n not in existing_cols and n not in BASE_COLS)
    for n in new_cols:
        header.append(n)
        ws.cell(row=1, column=len(header)).value = n

    col_index = {name: i + 1 for i, name in enumerate(header)}  # 1-based

    for orig_idx in added_indices:
        s = new_songs[orig_idx]
        votes = song_votes_by_added.get(orig_idx, [])
        up = sum(1 for _, p in votes if p > 0)
        down = sum(1 for _, p in votes if p < 0)
        row_vals = [None] * len(header)
        row_vals[0] = s.get("league", "")
        row_vals[1] = s.get("round", "")
        row_vals[2] = s.get("roundNum", 0)
        row_vals[3] = s.get("song", "")
        row_vals[4] = s.get("artist", "")
        row_vals[5] = s.get("album", "")
        row_vals[6] = s.get("score", 0)
        row_vals[7] = up
        row_vals[8] = down
        row_vals[9] = s.get("place", "")
        row_vals[10] = s.get("submitter", "")
        for name, pts in votes:
            ci = col_index.get(name)
            if ci:
                row_vals[ci - 1] = pts
        ws.append(row_vals)

    wb.save(str(VOTES_XLSX))
    print(f"{VOTES_XLSX.name} updated: {len(added_indices)} new rows"
          + (f", {len(new_cols)} new player columns" if new_cols else ""))


def _save_analytics(analytics, added_indices=None):
    """Merge new analytics data into rtml_analytics.json.

    Unlike a blind dict.update(), each key is merged conservatively so a
    caller that only supplies this week's new data can never wipe out
    history:
      - players: union, order-preserving, append-only
      - playerStats: merged by 'name' (update existing / append new)
      - roundSummaries: merged by (league, round) (update existing / append new)
      - songVotes: APPEND-ONLY. If added_indices is given and its length
        matches the supplied songVotes list, only the vote arrays for songs
        that were actually added (post-dedup) get appended, in order. This
        keeps songVotes positionally aligned with the SONGS/xlsx row order.
        Falls back to appending any tail entries beyond the current length
        if the caller already passed a merged array; never shrinks or
        replaces existing entries.
    """
    existing = {}
    if ANALYTICS_JSON.exists():
        existing = json.loads(ANALYTICS_JSON.read_text(encoding='utf-8'))

    if 'players' in analytics:
        merged_players = list(existing.get('players', []))
        seen = set(merged_players)
        for p in analytics['players']:
            if p not in seen:
                merged_players.append(p)
                seen.add(p)
        existing['players'] = merged_players

    if 'playerStats' in analytics:
        by_name = {p.get('name'): p for p in existing.get('playerStats', [])}
        for p in analytics['playerStats']:
            by_name[p.get('name')] = p
        existing['playerStats'] = list(by_name.values())

    if 'roundSummaries' in analytics:
        by_key = {(r.get('league'), r.get('round')): r for r in existing.get('roundSummaries', [])}
        for r in analytics['roundSummaries']:
            by_key[(r.get('league'), r.get('round'))] = r
        existing['roundSummaries'] = list(by_key.values())

    if 'songVotes' in analytics:
        existing_sv = list(existing.get('songVotes', []))
        new_sv = analytics['songVotes']
        if added_indices is not None and len(new_sv) == len(added_indices):
            # Caller supplied exactly one vote-array per song that was
            # actually added (post-dedup) — append directly, in order.
            existing_sv = existing_sv + new_sv
        elif len(existing_sv) == 0:
            existing_sv = new_sv
        elif len(new_sv) > len(existing_sv):
            # Caller passed an existing+new merged array — append only the tail
            existing_sv = existing_sv + new_sv[len(existing_sv):]
        else:
            print(f"WARNING: songVotes merge ambiguous (existing={len(existing_sv)}, "
                  f"supplied={len(new_sv)}) — keeping existing data untouched to avoid loss")
        existing['songVotes'] = existing_sv

    ANALYTICS_JSON.write_text(json.dumps(existing, ensure_ascii=False), encoding='utf-8')
    print(f"Analytics JSON updated (merged, not replaced)")


def _build_songs_js(songs):
    """Build the compact JS SONGS array string from song list."""
    js_lines = []
    for s in songs:
        def esc(v):
            if not v: return '""'
            return json.dumps(str(v))
        js_lines.append(
            f'[{esc(s["league"])},{esc(s["round"])},{s["roundNum"]},'
            f'{esc(s["song"])},{esc(s["artist"])},{esc(s["album"])},'
            f'{esc(s["place"])},{esc(s["submitter"])}]'
        )
    return "const SONGS=[" + ",".join(js_lines) + "];"


def regenerate_html(season_complete=None):
    """Re-embed the full song database and analytics into both HTML files.

    season_complete: optional league name to embed as SEASON_FINALE (see
    update_from_list docstring)."""
    import openpyxl
    wb = openpyxl.load_workbook(str(XLSX))
    ws = wb.active

    songs = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        league, round_name, round_num, song, artist, album, score, place, submitter = row
        songs.append({
            "league": league or "",
            "round": round_name or "",
            "roundNum": round_num or 0,
            "song": song or "",
            "artist": artist or "",
            "album": album or "",
            "score": score or 0,
            "place": place or "",
            "submitter": submitter or ""
        })

    new_data = _build_songs_js(songs)

    # Stats
    unique_rounds = len(set((s["league"], s["round"]) for s in songs))
    unique_players = len(set(s["submitter"] for s in songs if s["submitter"]))
    seasons = len(set(s["league"] for s in songs))
    total = len(songs)

    # ── Regenerate music_league.html + index.html ──────────────────────────
    html = HTML.read_text(encoding="utf-8")
    html = _splice_block(html, "const SONGS=", new_data)
    html = re.sub(r"countUp\(document\.getElementById\('stat-songs'\),\d+,",
                  f"countUp(document.getElementById('stat-songs'),{total},", html)
    html = re.sub(r"countUp\(document\.getElementById\('stat-rounds'\),\d+,",
                  f"countUp(document.getElementById('stat-rounds'),{unique_rounds},", html)
    html = re.sub(r"countUp\(document\.getElementById\('stat-seasons'\),\d+,",
                  f"countUp(document.getElementById('stat-seasons'),{seasons},", html)
    html = re.sub(r"countUp\(document\.getElementById\('stat-players'\),\d+,",
                  f"countUp(document.getElementById('stat-players'),{unique_players},", html)
    _write_validated(HTML, html, "music_league.html")
    print(f"music_league.html regenerated: {total} songs, {unique_rounds} rounds, {seasons} seasons, {unique_players} players")

    # ── Regenerate music_league_votes.html ─────────────────────────────────
    _regenerate_votes_html(songs, total, unique_rounds, seasons, unique_players, new_data, season_complete)


_TEMP_BANNER_RE = re.compile(
    r"<!-- TEMP-BANNER:([A-Za-z0-9_-]+):START expires=([0-9T:+-]+) -->.*?"
    r"<!-- TEMP-BANNER:\1:END -->\n?",
    re.DOTALL,
)


def _strip_expired_temp_banners(html, label=""):
    """Remove any temporary popup-banner block (e.g. a "Welcome back to
    Season N!" banner) whose embedded expiry timestamp has passed.

    Banners are wrapped like:
      <!-- TEMP-BANNER:<id>:START expires=<ISO8601> --> ... <!-- TEMP-BANNER:<id>:END -->
    so a banner added for one week's announcement is automatically cleaned
    out of the HTML the next time this script regenerates it, even if
    nobody remembers to remove it by hand. Banners not yet expired are left
    in place untouched."""
    import datetime

    def _maybe_strip(m):
        banner_id, expires_str = m.group(1), m.group(2)
        try:
            expires = datetime.datetime.fromisoformat(expires_str)
            if expires.tzinfo is None:
                expires = expires.replace(tzinfo=datetime.timezone.utc)
            now = datetime.datetime.now(datetime.timezone.utc)
            if now >= expires:
                print(f"{label}: removing expired temp banner '{banner_id}' (expired {expires_str})")
                return ""
        except ValueError:
            pass
        return m.group(0)

    return _TEMP_BANNER_RE.sub(_maybe_strip, html)


def _regenerate_votes_html(songs, total, unique_rounds, seasons, unique_players, new_data, season_complete=None):
    """Update music_league_votes.html with new SONGS, stats, and analytics."""
    if not VOTES_HTML.exists():
        print("music_league_votes.html not found, skipping")
        return

    html = VOTES_HTML.read_text(encoding="utf-8")
    html = _strip_expired_temp_banners(html, "music_league_votes.html")

    # Replace SONGS array
    html = _splice_block(html, "const SONGS=", new_data)

    # Update countUp stats
    html = re.sub(r"countUp\(document\.getElementById\('stat-songs'\),\d+,",
                  f"countUp(document.getElementById('stat-songs'),{total},", html)
    html = re.sub(r"countUp\(document\.getElementById\('stat-rounds'\),\d+,",
                  f"countUp(document.getElementById('stat-rounds'),{unique_rounds},", html)
    html = re.sub(r"countUp\(document\.getElementById\('stat-seasons'\),\d+,",
                  f"countUp(document.getElementById('stat-seasons'),{seasons},", html)
    html = re.sub(r"countUp\(document\.getElementById\('stat-players'\),\d+,",
                  f"countUp(document.getElementById('stat-players'),{unique_players},", html)

    # Update analytics bar pills
    html = re.sub(r'<span class="abar-pill">\d+ Seasons</span>',
                  f'<span class="abar-pill">{seasons} Seasons</span>', html)
    html = re.sub(r'<span class="abar-pill">\d+ Rounds</span>',
                  f'<span class="abar-pill">{unique_rounds} Rounds</span>', html)
    html = re.sub(r'<span class="abar-pill">[\d,]+ Songs</span>',
                  f'<span class="abar-pill">{total:,} Songs</span>', html)

    # Update result-count span (static fallback)
    html = re.sub(r'<span id="result-count">[\d,]+</span>',
                  f'<span id="result-count">{total:,}</span>', html)

    # If analytics JSON has updated data, re-embed PLAYERS and PLAYER_STATS
    if ANALYTICS_JSON.exists():
        analytics = json.loads(ANALYTICS_JSON.read_text(encoding="utf-8"))

        if analytics.get("players"):
            players_js = "const PLAYERS = " + json.dumps(analytics["players"], ensure_ascii=False) + ";"
            html = _splice_block(html, "const PLAYERS = ", players_js)

        if analytics.get("playerStats"):
            stats_js = "const PLAYER_STATS = " + json.dumps(analytics["playerStats"], ensure_ascii=False) + ";"
            html = _splice_block(html, "const PLAYER_STATS = ", stats_js)

        if analytics.get("roundSummaries"):
            rs_js = "const ROUND_SUMMARIES = " + json.dumps(analytics["roundSummaries"], ensure_ascii=False) + ";"
            html = _splice_block(html, "const ROUND_SUMMARIES = ", rs_js)

        if analytics.get("songVotes"):
            sv_js = "const SONG_VOTES = " + json.dumps(analytics["songVotes"], ensure_ascii=False) + ";"
            html = _splice_block(html, "const SONG_VOTES = ", sv_js)

    # ── Hall of Fame: auto-update from xlsx scores ─────────────────────
    # HOF_OFFICIAL, HOF_IN_PROGRESS, SEASON_ORDER, SEASON_LABELS are all
    # hardcoded in the HTML. We recompute them from the full songs list so
    # the current season's standings update automatically each week without
    # manual edits.
    hof = _compute_hof_data(songs)

    hof_official_js = ("const HOF_OFFICIAL = "
                       + json.dumps(hof["hof_official"], ensure_ascii=False)
                       + ";")
    html = _splice_block(html, "const HOF_OFFICIAL = ", hof_official_js,
                         open_char='{', close_char='}')

    season_order_js = ("const SEASON_ORDER = "
                       + json.dumps(hof["season_order"], ensure_ascii=False)
                       + ";")
    html = _splice_block(html, "const SEASON_ORDER = ", season_order_js)

    season_labels_js = ("const SEASON_LABELS = "
                        + json.dumps(hof["season_labels"], ensure_ascii=False)
                        + ";")
    html = _splice_block(html, "const SEASON_LABELS = ", season_labels_js,
                         open_char='{', close_char='}')

    in_progress_js = ('const HOF_IN_PROGRESS = '
                      + json.dumps(hof["in_progress"])
                      + ';')
    html = re.sub(r'const HOF_IN_PROGRESS = "[^"]*";',
                  in_progress_js, html)

    print(f"Hall of Fame updated: in-progress season = {hof['in_progress']}, "
          f"current top 3 = {hof['hof_official'].get(hof['in_progress'], [])}")

    # ── Season finale banner ────────────────────────────────────────────
    # Only touch SEASON_FINALE when explicitly told a season just wrapped;
    # otherwise leave whatever's already embedded (so the banner for last
    # week's finale doesn't get silently cleared by an unrelated update).
    if season_complete:
        finale = _compute_season_finale(hof, season_complete)
        if finale:
            finale_js = "const SEASON_FINALE = " + json.dumps(finale, ensure_ascii=False) + ";"
            html = _splice_block(html, "const SEASON_FINALE = ", finale_js,
                                 open_char='{', close_char='}')
            print(f"Season finale banner set: {finale['league']} winner = "
                  f"{finale['winner']} ({finale['score']} pts)")
        else:
            print(f"WARNING: --season-complete {season_complete!r} has no HOF data, banner not set")

    _write_validated(VOTES_HTML, html, "music_league_votes.html")
    # index.html served by GitHub Pages — keep it byte-identical to the votes site
    _write_validated(BASE / "index.html", html, "index.html")
    print(f"music_league_votes.html + index.html regenerated: {total} songs, {unique_rounds} rounds, {seasons} seasons, {unique_players} players")


def _cli_main():
    if '--validate' in sys.argv:
        # Standalone check: confirm the files currently on disk are
        # structurally sound, without regenerating anything. Useful as a
        # pre-push gate (e.g. run this right before uploading to GitHub).
        ok = True
        for path in (HTML, VOTES_HTML, BASE / "index.html"):
            if not path.exists():
                continue
            problems = validate_html(path.read_text(encoding="utf-8"), path.name)
            if problems:
                ok = False
                print(f"FAIL {path.name}:")
                for p in problems:
                    print(f"  - {p}")
            else:
                print(f"OK   {path.name}")
        sys.exit(0 if ok else 1)
    elif '--new-songs' in sys.argv:
        idx = sys.argv.index('--new-songs')
        data = json.loads(sys.argv[idx+1])
        analytics = None
        if '--analytics' in sys.argv:
            aidx = sys.argv.index('--analytics')
            analytics = json.loads(sys.argv[aidx+1])
        season_complete = None
        if '--season-complete' in sys.argv:
            sidx = sys.argv.index('--season-complete')
            season_complete = sys.argv[sidx+1]
        update_from_list(data, analytics, season_complete=season_complete)
    elif '--regenerate' in sys.argv:
        season_complete = None
        if '--season-complete' in sys.argv:
            sidx = sys.argv.index('--season-complete')
            season_complete = sys.argv[sidx+1]
        regenerate_html(season_complete=season_complete)
    else:
        print("Usage:")
        print("  python3 update_music_league.py --new-songs '[...]' --analytics '{...}' [--season-complete 'League Name']")
        print("  python3 update_music_league.py --regenerate [--season-complete 'League Name']")
        print("  python3 update_music_league.py --validate")


if __name__ == "__main__":
    try:
        _cli_main()
    except (BlockNotFound, RuntimeError) as e:
        print(f"\nABORTED — no files were left in a broken state: {e}", file=sys.stderr)
        sys.exit(1)
