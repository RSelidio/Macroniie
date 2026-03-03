# Macronnie - Macro Recorder & Player

A lightweight Windows macro recorder/player for automating mouse and keyboard inputs with global hotkey controls.

**Owner:** Ronnie

##  Features

- **Record & Playback**: Captures mouse movements, clicks, scrolls, and keyboard input
- **Custom Global Hotkeys**: Set your own Record/Stop/Play shortcuts in the UI
- **Loop Support**: Repeat macros continuously or for a specific number of iterations
- **Loop Timer**: Set delay (seconds) between each playback cycle
- **Loop Counter**: Specify exactly how many times to repeat (e.g., run macro 5 times then stop)
- **Background Keys**: Configure up to 10 independent keyboard keys to press at custom intervals during playback (e.g., F1 every 30s)
- **DPI-Aware**: Handles multiple monitor setups and DPI scaling
- **Standalone Executable**: Pre-built .exe available (no Python required for end-users)
- **Simple UI**: Intuitive interface showing event count, duration, and hotkey instructions

##  Hotkeys

| Key | Action |
|-----|--------|
| **Custom (UI setting)** | Start recording |
| **Custom (UI setting)** | Stop recording |
| **Custom (UI setting)** | Play macro |
| **Custom (UI setting)** | Kill (stop recording/playback) |

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
   - Press your configured **Record** hotkey to start recording
   - Move mouse, click, scroll, and type on your keyboard
   - Press your configured **Stop** hotkey (or Ctrl+Esc) to stop
   - All events are timestamped and saved in memory

2. **Playback Phase**:
   - Press your configured **Play** hotkey to play back the recorded macro
   - Mouse movements and clicks are replayed with precise timing
   - Keyboard inputs are re-entered in sequence
   - Check "Loop playback" to enable looping
   - Set **Loop interval (s)** to add a delay between repetitions (e.g., 0 = no delay, 30 = 30 seconds between runs)
   - Set **Loop Times** to specify exact repetitions (e.g., 1 = single run, 5 = run 5 times then stop). Leave at 1 for single playback.

3. **Data Format**:
   - Events stored as JSON with timestamps (in seconds)
   - Event types: `move`, `click`, `scroll`, `key`
   - Playback uses relative timing for accurate reproduction

4. **Background Keys** (automation during playback):
   - Check the checkbox for each key slot you want to use (up to 10 total)
   - Enter key name: `f1`–`f12` for function keys, or single letters like `a`, `space`, etc.
   - Set interval in seconds: how often this key should press (e.g., 30 = every 30 seconds)
   - Keys activate when you press Play and run independent of macro events
   - Example: F1 every 30s + F2 every 60s will run side-by-side during playback

##  Limitations & Notes

- **Games with Anti-Cheat**: Some games may not respond to simulated inputs
- **Admin Privileges**: Running as administrator improves compatibility
- **Raw Input Games**: Some games (FPS, etc.) may not register macro inputs
- **Screen Capture**: Position-based macros may need adjustment after resolution/DPI changes

##  Quick Start Example

1. Open a game or application
2. Press your configured **Record** hotkey (macro recorder starts)
3. Perform your actions (move mouse, click, type)
4. Press your configured **Stop** hotkey (recording stops)
5. Press your configured **Play** hotkey (playback begins)
6. Check "Loop" and press your configured **Play** hotkey to repeat

##  License

Open source for personal use.

---

**Made by Ronnie** | [GitHub](https://github.com/RSelidio/Macroniie)
