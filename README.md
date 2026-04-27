DiGiCo to Reaper Converter v3.2
================================

INSTALLATION:
1. Drag "DiGiCo to Reaper.app" to your Applications folder
2. Double-click to launch
3. Look for the 🎛️ icon in your menu bar (top-right of screen)

USAGE:
1. Export a session report from your DiGiCo console (.rtf file)
   - Make sure to Include Channels
2. Click the 🎛️ menu bar icon and select "Open Converter"
3. Your browser will open automatically to the converter
4. Upload your .rtf file using the upload area, or add channels manually
5. Select/deselect channels and organize as needed
6. Choose your stereo mode and click Download
7. Import into Reaper:
   - Track → Insert tracks from template → Select file
   - OR drag the .RTrackTemplate file directly into Reaper

SESSIONS (TABS):
The converter supports multiple sessions open at once using tabs.
- Click "+" to add a new session tab
- Double-click a tab name to rename it
- Click the × on a tab to close it
- Each tab is fully independent — channels, colors, and selections do not carry over between tabs

UPLOADING FILES:
- Drag and drop an .rtf file onto the upload area, or click to browse
- To load a new file into an existing session, click "Upload New File" — you will be prompted before the session is cleared
- If no file is loaded, you can still build a session by adding channels manually

ADDING AND EDITING CHANNELS:
- Click "+ Add Channel Manually" to add channels that aren't in your session report
- Choose the channel type (Input, Aux, Group, Matrix), name, quantity, and whether it's stereo
- Double-click any channel name to rename it inline
- Use the Up/Down arrow keys while renaming to move to the next or previous channel
- Click the ✕ on any channel row to remove it immediately
- Select one or more channels and press Delete (or Backspace on Windows) to remove with confirmation

SELECTING CHANNELS:
- Click any channel row to highlight it
- Cmd/Ctrl+Click to toggle individual channels on/off
- Shift+Click to select a range
- Cmd/Ctrl+A to select all
- Click empty space in the list to deselect all

REORDERING CHANNELS:
- Drag any channel row to reorder it
- Drag a highlighted channel to move all highlighted channels together as a group
- The list will scroll automatically when dragging near the top or bottom

QUICK SELECTIONS:
Each section (Inputs, Aux Outputs, Group Outputs, Matrix Outputs) has a checkbox to quickly include or exclude the entire section from the export.

CHANNEL COLORS:
- Click the color dot on any channel row to assign a color that will appear on the track in Reaper
- Right-click the dot (or click ✕) to clear the color
- Click the color dot next to a section label to apply one color to every channel in that section
- When multiple channels are highlighted, a bulk toolbar appears to apply or clear colors across all selected channels at once
- Stereo channels (split into L/R) will both receive the same color

STEREO MODE:
Use the stereo toggle to choose how stereo channels are exported:
- Split into Mono: Each stereo channel becomes two separate mono tracks (L and R)
- Keep as Stereo: Each stereo channel becomes one stereo track

REAPER TRACK SETTINGS:
All exported tracks are:
- Record armed and ready to receive input
- Routed sequentially to hardware inputs (mono channels take one input, stereo channels take two)
- Colored according to any colors you assigned

EXPORT OPTIONS:
- Download .RTrackTemplate — imports directly into Reaper as a track template
- Export as CSV — saves the channel list as a .csv file for use in spreadsheets or other tools

UNDO:
- Click the Undo button (or use your browser's standard undo if available) to step back through changes
- Up to 50 undo steps are stored per session

DARK MODE:
Click the "Dark Mode" button in the top-right corner to switch the interface to a dark background.

MENU BAR OPTIONS:
- Open Converter — opens the converter in your default browser
- Restart Server — restarts the local server if something goes wrong
- Quit — properly closes the app and stops the server

LUA SCRIPT (OPTIONAL — REAPER DIRECT IMPORT):
Also included is DiGiCo_to_Reaper.lua, a ReaScript you can run directly inside
Reaper to import input channels from a DiGiCo session report without using the
converter app or browser at all.

What it does:
- Prompts you to select a DiGiCo .rtf session report
- Parses all Input Channels from the report
- Creates one track per channel (stereo channels split into L/R mono tracks)
- Assigns sequential hardware inputs
- Colors all tracks gray (#8E8E93) by default

Note: The Lua script only imports Input Channels. It does not support Aux,
Group, or Matrix outputs, custom colors, stereo-keep mode, or CSV export.
For full control over your session, use the converter app instead.

SETTING UP THE LUA SCRIPT IN REAPER:
1. In Reaper, go to Actions > Load ReaScript
2. Browse to and select "DiGiCo_to_Reaper.lua"
3. Reaper will load the script and add it to your Actions list
4. To run it: Actions > Action List, find "DiGiCo_to_Reaper", click Run
5. (Optional) To assign a keyboard shortcut:
   - Find the script in the Action List
   - Click "Add shortcut" and press your preferred key combination
     (e.g. Cmd+Shift+I on Mac)

TROUBLESHOOTING:
- If the browser doesn't open automatically, click "Open Converter" from the menu bar icon
- The port number is shown in the menu bar tooltip (usually :8081)
- If the app won't open on first launch, right-click → Open to bypass Gatekeeper
- If the browser shows a disconnect message, the app has been closed — relaunch from Applications

CHANNEL TYPES:
- Inputs — input channels from the console
- Aux Outputs — aux/monitor mixes
- Group Outputs — group/subgroup buses
- Matrix Outputs — matrix outputs
- Custom — channels added manually

Built by: Michael Leckrone
Contact: leckroneaudio@gmail.com
Version: 3.2
