# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This repository contains two independent mini-projects:

1. **Pomodoro Timer** — A standalone single-file web app (`pomodoro-timer.html`)
2. **PPT Generator** — A Node.js script that generates a PowerPoint presentation from Excel data (`generate_ppt.js`)

No build step, bundler, or framework is used. There are no tests.

## Commands

```bash
# Generate the PPT (reads from Desktop Excel file, writes PPT to Desktop)
node generate_ppt.js

# Install dependencies (only needed for generate_ppt.js)
npm install
```

The Pomodoro timer has no dependencies — just open `pomodoro-timer.html` in a browser.

## Architecture

### Pomodoro Timer (`pomodoro-timer.html`)

A fully self-contained single HTML file (CSS in `<style>`, JS in `<script>`). No external resources.

- **State management**: A plain `state` object tracks timer state (`isRunning`, `isBreak`, `isPaused`, durations, session counts). Session stats persist to `localStorage` under keys `pomo_sessions`, `pomo_total_min`, `pomo_today`.
- **Timer**: `setInterval`-based countdown (`tick()` every 1s). When a session finishes, it automatically toggles between work and break modes and starts the next session.
- **Audio**: Uses the Web Audio API (`AudioContext` + `OscillatorNode`) for beep/tick/finish sounds — no audio files.
- **Notifications**: Uses the Web Notifications API for desktop alerts.
- **Visual**: SVG ring progress bar (753.98 circumference = 2π × 120 radius). CSS custom properties on `:root` define the color scheme (dark theme, red accent for work, green accent for breaks).
- **Keyboard shortcuts**: Space (start/pause), R (reset), S (skip).

### PPT Generator (`generate_ppt.js`)

Reads an Excel file from the user's Desktop, parses structured markdown content from a specific row, and builds a themed PowerPoint presentation.

- **Data source**: `~/Desktop/create_recrod_brazil6031659.xlsx` — filtered to row where `countryid === 'japan06032028'` and `number === '1'`.
- **Content parsing**: Splits the `information` column by newlines, identifies `#### ` headings as subsection titles, extracts bullet points (filters out short/empty lines, strips `【来源：...】` and `【发布时间：...】` metadata).
- **Slide structure**: Cover → Agenda → 3 sections (each: divider → content slides → takeaways) → Summary → Thank You. Each content slide uses a 1- or 2-column layout depending on point count.
- **Styling**: Uses a consistent color palette (navy, blue, teal, red, gold) with `pptxgenjs` shapes for layout. Fonts: Microsoft YaHei for Chinese text, Arial for English labels.
- **Output**: Writes to `~/Desktop/日本国情概况.pptx`.
