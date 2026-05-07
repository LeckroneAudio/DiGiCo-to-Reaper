-- DiGiCo to Reaper — Input Channel Importer
-- Parses a DiGiCo RTF session report and creates tracks for all Input Channels.
-- Stereo channels (suffix 's') are split into two mono tracks (Name L / Name R).
--
-- Installation:
--   1. Actions > Load ReaScript, select this file
--   2. Assign shortcut: Actions > Action List → find script → Add shortcut → Cmd+Shift+I

-- ============================================================
-- UTILITIES
-- ============================================================

local function split_str(str, sep)
    local result, i, sep_len = {}, 1, #sep
    while true do
        local j = str:find(sep, i, true)
        if j then
            result[#result + 1] = str:sub(i, j - 1)
            i = j + sep_len
        else
            result[#result + 1] = str:sub(i)
            break
        end
    end
    return result
end

local function trim(s)
    return (s:gsub("^%s+", ""):gsub("%s+$", ""))
end

local function clean_rtf_field(s)
    s = s:gsub("\\'[%x][%x]", "")
    s = s:gsub("\\[%a]+%-?[%d]*%s?", "")
    s = s:gsub("\\[^%a%d%s]", "")
    s = s:gsub("[{}]", "")
    s = s:gsub("%s+", " ")
    return trim(s)
end

-- ============================================================
-- PARSER — INPUT CHANNELS ONLY
-- ============================================================

local function parse_input_channels(content)
    local channels = {}
    local lines = split_str(content, "\\par")
    local in_section, found_header = false, false

    for _, line in ipairs(lines) do
        if line:find("Input Channels", 1, true) and line:find("\\b", 1, true) then
            in_section, found_header = true, false
        elseif in_section and (
            line:find("Aux Outputs",    1, true) or
            line:find("Group Outputs",  1, true) or
            line:find("Matrix Outputs", 1, true) or
            line:find("Matrix Inputs",  1, true) or
            line:find("Control Groups", 1, true) or
            line:find("VCA Groups",     1, true)
        ) then
            break  -- done with input channels
        end

        if in_section and not found_header then
            if line:lower():find("name", 1, true) then found_header = true end
            goto continue
        end

        if in_section and found_header then
            local parts = split_str(line, "\\tab")
            if #parts >= 2 then
                local num  = clean_rtf_field(parts[1])
                local name = clean_rtf_field(parts[2])
                local ch   = num:match("^(%d+s?)$")

                if ch and name ~= "" and name:lower() ~= "name" then
                    local dup = false
                    for _, existing in ipairs(channels) do
                        if existing.number == ch then dup = true; break end
                    end
                    if not dup then
                        channels[#channels + 1] = { number=ch, name=name }
                    end
                end
            end
        end

        ::continue::
    end

    return channels
end

-- ============================================================
-- TRACK CREATION
-- ============================================================

local INPUT_COLOR = { r=142, g=142, b=147 }  -- #8E8E93 gray

local function create_tracks(channels)
    local proj      = 0
    local n         = 0
    local first_idx = reaper.CountTracks(proj)

    reaper.Undo_BeginBlock()
    reaper.PreventUIRefresh(1)

    for _, ch in ipairs(channels) do
        local is_stereo = ch.number:sub(-1) == "s"
        local color     = reaper.ColorToNative(INPUT_COLOR.r, INPUT_COLOR.g, INPUT_COLOR.b)
        local names     = is_stereo and { ch.name.." L", ch.name.." R" } or { ch.name }

        for _, nm in ipairs(names) do
            local idx = reaper.CountTracks(proj)
            reaper.InsertTrackAtIndex(idx, true)
            local tr = reaper.GetTrack(proj, idx)
            reaper.GetSetMediaTrackInfo_String(tr, "P_NAME", nm, true)
            reaper.SetTrackColor(tr, color)
            n = n + 1
        end
    end

    -- Sequential hardware input assignment (I_RECINPUT = track's 0-based index)
    for i = first_idx, reaper.CountTracks(proj) - 1 do
        reaper.SetMediaTrackInfo_Value(reaper.GetTrack(proj, i), "I_RECINPUT", i)
    end

    reaper.TrackList_AdjustWindows(false)
    reaper.UpdateArrange()
    reaper.PreventUIRefresh(-1)
    reaper.Undo_EndBlock("DiGiCo: Import input channels", -1)
    return n
end

-- ============================================================
-- MAIN
-- ============================================================

local ok, filepath = reaper.GetUserFileNameForRead("", "Select DiGiCo Session Report (.rtf)", "rtf")
if not ok then return end

local f, err = io.open(filepath, "r")
if not f then
    reaper.ShowMessageBox("Could not open file:\n" .. tostring(err), "DiGiCo to Reaper", 0)
    return
end
local content = f:read("*all"); f:close()

if not content or content == "" then
    reaper.ShowMessageBox("File is empty or could not be read.", "DiGiCo to Reaper", 0)
    return
end

local channels = parse_input_channels(content)

if #channels == 0 then
    reaper.ShowMessageBox(
        "No input channels found.\n\nMake sure this is a DiGiCo RTF session report\n(File > Print Session Report on the console).",
        "DiGiCo to Reaper", 0)
    return
end

create_tracks(channels)
