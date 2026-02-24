# Macronnie - Macro Recorder & Player

A lightweight Windows macro recorder/player for automating mouse and keyboard inputs with global hotkey controls.

**Owner:** Ronnie

##  Features

- **Record & Playback**: Captures mouse movements, clicks, scrolls, and keyboard input
- **Global Hotkeys**: Control recording/playback from anywhere using Ctrl+1/2/3/Esc
- **Loop Support**: Repeat macros continuously until stopped
- **DPI-Aware**: Handles multiple monitor setups and DPI scaling
- **Standalone Executable**: Pre-built .exe available (no Python required for end-users)
- **Simple UI**: Intuitive interface showing event count, duration, and hotkey instructions

##  Hotkeys

| Key | Action |
|-----|--------|
| **Ctrl+1** | Start recording |
| **Ctrl+2** | Stop recording |
| **Ctrl+3** | Play macro |
| **Ctrl+Esc** | Kill (stop recording or playback) |

##  Installation & Usage

### Option 1: Pre-Built Executable (Easiest)

1. Download `Macronnie.zip` from [Releases](https://github.com/RSelidio/Macroniie/releases)
2. Extract the ZIP file
3. Double-click `Macronnie.exe`
4. No Python or dependencies needed!

### Option 2: Run from Source

#### Requirements
- Python 3.10+
- Windows OS

#### Setup

`powershell
git clone https://github.com/RSelidio/Macroniie.git
cd Macroniie

py -m venv .venv
.\.venv\Scripts\Activate.ps1

py -m pip install -r requirements.txt
`

#### Run

`powershell
py main.py
`

##  Dependencies

- pynput>=1.7.7 - Global keyboard/mouse input capture
- psutil>=5.9.8 - Process and window detection

##  How It Works

1. **Recording Phase**:
   - Press **Ctrl+1** to start recording
   - Move mouse, click, scroll, and type on your keyboard
   - Press **Ctrl+2** (or Ctrl+Esc) to stop
   - All events are timestamped and saved in memory

2. **Playback Phase**:
   - Press **Ctrl+3** to play back the recorded macro
   - Mouse movements and clicks are replayed with precise timing
   - Keyboard inputs are re-entered in sequence
   - Check the "Loop" checkbox to repeat continuously

3. **Data Format**:
   - Events stored as JSON with timestamps (in seconds)
   - Event types: `move`, `click`, `scroll`, `key`
   - Playback uses relative timing for accurate reproduction

##  Limitations & Notes

- **Games with Anti-Cheat**: Some games may not respond to simulated inputs
- **Admin Privileges**: Running as administrator improves compatibility
- **Raw Input Games**: Some games (FPS, etc.) may not register macro inputs
- **Screen Capture**: Position-based macros may need adjustment after resolution/DPI changes

##  Quick Start Example

1. Open a game or application
2. Press **Ctrl+1** (macro recorder starts)
3. Perform your actions (move mouse, click, type)
4. Press **Ctrl+2** (recording stops)
5. Press **Ctrl+3** (playback begins)
6. Check "Loop" and press **Ctrl+3** to repeat

##  License

Open source for personal use.

---

**Made by Ronnie** | [GitHub](https://github.com/RSelidio/Macroniie)
