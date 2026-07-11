# 🌲 TripVisualizer — Black Forest Family Trip 2026

An interactive HTML trip guide and Excel schedule generator for a family vacation in the Black Forest, Germany (July 15–25, 2026).

---

## 📁 Files

| File | Description |
|------|-------------|
| `trip_guide.html` | Interactive map & day-by-day sidebar — the main guide |
| `generate_trip.py` | Python script to generate a color-coded Excel schedule |
| `trip_knowledge.md` | Full knowledge base: logistics, itinerary, coordinates, tips |

---

## 🗺️ trip_guide.html — Interactive Map Guide

Open directly in any modern browser. No server required.

### Features

#### Interactive Leaflet.js Map
- **21 locations** plotted across the Black Forest and Rhine region
- Color-coded markers by activity type (nature, lakes, attractions, cities, hotels, flights)
- **Two map layers**: OpenStreetMap (street) and OpenTopoMap (topographic) — switchable via layer control
- **Route lines** drawn when a day is selected, connecting the day's stops in order
- Clicking a marker opens a popup with Hebrew description and link to the official website

#### Day-by-Day Sidebar
- **11-day tabs** (Jul 15–25) + "All Days" overview
- Each day shows:
  - **Hotel card** — the night's accommodation, listed as the first item before activities; click to fly the map to the hotel
  - **Activity cards** — time, duration, Hebrew description, and note for each stop
  - **Driving distance badges** — validated km shown on all drive activities (e.g. `🛣 ~80 ק"מ`)
  - **Weather badges** — per-activity weather icon + temperature (online mode only)
- "All Days" view shows a summary card per day with title and highlights

#### 🌐 Online / Offline Mode
Toggle the **⚡ מצב אונליין** button in the header to switch modes.

**Online mode fetches live data from [Open-Meteo](https://open-meteo.com) (free, no API key required):**
- 11-day weather forecast for the entire trip
- Floating weather panel (bottom-left of map) with scrollable day cards showing:
  - Weather icon + Hebrew description
  - Max / min temperature
  - Precipitation in mm
  - Click any day card to jump to that day's itinerary
- Per-activity weather badges injected inline into activity cards
- Weather data is cached per 0.1° grid to minimize API calls
- Panel is collapsible (▼/▲ toggle)

#### Activity Types & Colors

| Type | Color | Examples |
|------|-------|---------|
| 🌲 Nature / Hiking | Green | Ravennaschlucht, Wutach Gorge, Feldberg |
| 🏊 Lakes / Swimming | Blue | Titisee, Schluchsee |
| 🎢 Attractions | Orange | Europa-Park, Hasenhorn Rodelbahn, Action Forest |
| 🏛️ Cities / Culture | Purple | Freiburg, Stein am Rhein, St. Blasien |
| 🏨 Hotels | Dark Blue | Hofgut Sternen, Waldshut-Tiengen |
| ✈️ Flights / Transport | Pink | ZRH arrivals/departures |
| 🚗 Driving | Gray | All road transfers |

---

## 📊 generate_trip.py — Excel Schedule Generator

Generates `Black_Forest_Trip_2026.xlsx` using the `openpyxl` library.

### Features
- **11 columns** (one per day, Jul 15–25), time rows from 07:00 to 23:00
- Color-coded cells by activity type matching the HTML guide palette
- Merged hotel header rows for each hotel block
- Hebrew RTL sheet view (`sheet_view.rightToLeft = True`)
- Legend row at the bottom

### Requirements
```bash
pip install openpyxl
python generate_trip.py
```

---

## 🗓️ Trip Overview

**Family:** Guy, עדי, עמית (16), נדב (14), איתי (11)

**Hotels:**
| | Dates | Location |
|--|-------|---------|
| Hotel Hofgut Sternen | Jul 15–21 (6 nights) | Breitnau, Black Forest — at the mouth of Ravennaschlucht gorge |
| Waldshut-Tiengen hotel | Jul 21–25 (4 nights) | Waldshut-Tiengen — on the Rhine, Swiss border |

**Flights:**
- Outbound: OS135 · Land Zurich ZRH · 11:00 Jul 15
- Return: OS146 · Depart Zurich ZRH · 20:50 Jul 25

### Day-by-Day Highlights

| Day | Date | Title | Highlights |
|-----|------|-------|-----------|
| 1 | Jul 15 | Arrival + Ravennaschlucht | Land ZRH → drive to hotel → gorge walk |
| 2 | Jul 16 | Titisee + Feldberg + Badeparadies | Lake swim · cable car to 1493m · water park |
| 3 | Jul 17 | Europa-Park 🎢 | Full day at Germany's biggest theme park |
| 4 | Jul 18 | Todtnau + Hasenhorn + Blackforestline | Waterfalls · mountain coaster · suspension bridge |
| 5 | Jul 19 | Schluchsee + Action Forest | Kayak/SUP · rope park with zip-lines |
| 6 | Jul 20 | Triberg + Vogtsbauernhof | Germany's highest waterfalls · open-air museum |
| 7 | Jul 21 | Move + Wutach Gorge + Tannenmühle | "Grand Canyon of Black Forest" · farm lunch |
| 8 | Jul 22 | Rhine Falls + Stein am Rhein 🇨🇭 | Europe's most powerful waterfall · medieval town |
| 9 | Jul 23 | Albsteig + St. Blasien | Wild circular trail · baroque cathedral |
| 10 | Jul 24 | Freiburg + Schauinslandbahn | Old town · one of Germany's longest cable cars |
| 11 | Jul 25 | Return | Morning walk · drive to ZRH · fly home |

---

## 🛠️ Tech Stack

- **HTML/CSS/JS** — single-file, no build step, no dependencies except CDN
- **[Leaflet.js](https://leafletjs.com/) v1.9.4** — interactive map
- **[OpenStreetMap](https://www.openstreetmap.org/)** + **[OpenTopoMap](https://opentopomap.org/)** — free map tiles
- **[Open-Meteo API](https://open-meteo.com/)** — free weather forecast, no API key
- **Python + [openpyxl](https://openpyxl.readthedocs.io/)** — Excel generation

---

## 🚀 Usage

1. Clone the repo
2. Open `trip_guide.html` in a browser
3. Click day tabs to explore the itinerary
4. Hit **⚡ מצב אונליין** for live weather forecast
5. To regenerate the Excel: `pip install openpyxl && python generate_trip.py`
