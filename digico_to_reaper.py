#!/usr/bin/env python3
"""
DiGiCo to Reaper Track Template Converter
Parses DiGiCo session reports and generates Reaper track templates
"""

from http.server import HTTPServer, BaseHTTPRequestHandler
import io
import json
import os
import re
import struct
import subprocess
import tarfile
import uuid
import socket
import threading
import zlib
from urllib.parse import parse_qs

try:
    import rumps
except ImportError:
    print("Installing rumps...")
    subprocess.run(['pip3', 'install', 'rumps', '--break-system-packages'], check=True)
    import rumps

def parse_digico_rtf(rtf_content):
    """Parse DiGiCo RTF session report and extract all channel sections"""
    
    # Decode if it's bytes
    if isinstance(rtf_content, bytes):
        rtf_content = rtf_content.decode('utf-8', errors='ignore')
    
    # Result structure with all sections
    result = {
        'inputs': [],
        'aux': [],
        'groups': [],
        'matrix': []
    }
    
    # Split into lines by \par
    lines = rtf_content.split('\\par')
    
    # State tracking
    current_section = None
    found_header = False
    
    for i, line in enumerate(lines):
        # Detect section headers
        if 'Input Channels' in line and '\\b' in line:
            current_section = 'inputs'
            found_header = False
            print(f"Found Input Channels section at line {i}")
            continue
        elif 'Aux Outputs' in line and '\\b' in line:
            current_section = 'aux'
            found_header = False
            print(f"Found Aux Outputs section at line {i}")
            continue
        elif 'Group Outputs' in line and '\\b' in line:
            current_section = 'groups'
            found_header = False
            print(f"Found Group Outputs section at line {i}")
            continue
        elif 'Matrix Outputs' in line and '\\b' in line:
            current_section = 'matrix'
            found_header = False
            print(f"Found Matrix Outputs section at line {i}")
            continue
        elif 'Matrix Inputs' in line or 'Control Groups' in line:
            # End of sections we care about
            current_section = None
            continue
        
        # Skip header line in each section
        if current_section and not found_header:
            if 'name' in line.lower():
                found_header = True
                print(f"Found {current_section} header line, skipping")
                continue
        
        # Parse channel lines
        if current_section and found_header:
            parts = line.split('\\tab')
            
            if len(parts) >= 2:
                # Get channel number - different formats for different sections
                # Inputs: "1", "1s"
                # Aux: "A1", "A1s"
                # Groups: "G1", "G1s"  
                # Matrix: "M1", "M1s"
                if current_section == 'inputs':
                    channel_match = re.search(r'^(\d+s?)\s*$', parts[0].strip())
                else:
                    # Aux, Groups, Matrix have letter prefix
                    channel_match = re.search(r'^([AGM]\d+s?)\s*$', parts[0].strip())
                
                if channel_match:
                    channel_num = channel_match.group(1).strip()
                    channel_name = parts[1].strip()
                    
                    # Clean up name
                    channel_name = re.sub(r'\s+', ' ', channel_name)
                    
                    if channel_name and channel_name.lower() not in ['name', '']:
                        result[current_section].append({
                            'number': channel_num,
                            'name': channel_name,
                            'type': current_section
                        })
                        if len(result[current_section]) <= 3:
                            print(f"  [{current_section}] {channel_num}: {channel_name}")
    
    # Print summary
    print(f"\n=== PARSE SUMMARY ===")
    print(f"Inputs: {len(result['inputs'])}")
    print(f"Aux: {len(result['aux'])}")
    print(f"Groups: {len(result['groups'])}")
    print(f"Matrix: {len(result['matrix'])}")

    return result


def parse_rivage_pm_show_file(file_content):
    """Parse a Yamaha Rivage PM .RIVAGEPM show file and extract channel sections.

    The file is a Yamaha MBDF (Multi-Block Data Format) container.  The mixing
    data lives in the first zlib-compressed block that contains the b'EN00/mix'
    marker.  Inside that block a binary schema section (COL0 / PR entries)
    precedes a raw data section whose layout is described by the schema.
    """

    if isinstance(file_content, str):
        file_content = file_content.encode('latin-1')

    # ── 1. Locate and decompress the mixing block ──────────────────────────
    raw = None
    for pos in range(0, len(file_content) - 2):
        if file_content[pos:pos+2] not in (b'\x78\x01', b'\x78\x9c', b'\x78\xda'):
            continue
        try:
            candidate = zlib.decompress(file_content[pos:pos+200000])
            if b'EN00/mix' in candidate[:256]:
                raw = candidate
                break
        except Exception:
            continue

    if raw is None:
        return {'inputs': [], 'aux': [], 'groups': [], 'matrix': []}

    # ── 2. Walk COL0/PR schema entries to find data-section start ──────────
    schema = []
    pos = 196  # skip file/block headers
    while pos < len(raw) - 32:
        if raw[pos:pos+4] == b'COL0':
            name_b = raw[pos+4:pos+28]
            name = name_b[:name_b.find(0) if 0 in name_b else 24].decode('ascii', errors='replace')
            vals = struct.unpack('<5I', raw[pos+28:pos+48])
            schema.append({'kind': 'COL0', 'name': name, 'v': vals})
            pos += 48
        elif raw[pos:pos+3] == b'PR ':
            pos += 32
        else:
            break

    data_start = pos

    # ── 3. Build a map of top-level channel sections ───────────────────────
    # For each section (InputChannel, Mix, …) we need:
    #   data_offset  – byte offset from data_start to this section's records
    #   rec_size     – bytes per record
    #   count        – number of records
    #   name_offset  – byte offset of the Name field within each record
    #                  (equals v[2] of the immediately-following COL0Label)
    TARGET_SECTIONS = ('InputChannel', 'Mix', 'Matrix', 'Stereo')
    sections = {}
    for i, entry in enumerate(schema):
        if entry['kind'] != 'COL0' or entry['name'] not in TARGET_SECTIONS:
            continue
        v = entry['v']
        if v[4] < 1:
            continue
        # Look ahead for the next COL0Label to get the name field offset
        name_offset = None
        for j in range(i + 1, min(i + 60, len(schema))):
            if schema[j]['kind'] == 'COL0' and schema[j]['name'] == 'Label':
                name_offset = schema[j]['v'][2]
                break
        if name_offset is None:
            continue
        if entry['name'] not in sections:  # keep first occurrence only
            sections[entry['name']] = {
                'data_offset': v[2],
                'rec_size':    v[3],
                'count':       v[4],
                'name_offset': name_offset,
            }

    # ── 4. Extract channel names ───────────────────────────────────────────
    def read_names(sect_name):
        info = sections.get(sect_name)
        if not info:
            return []
        sect_start = data_start + info['data_offset']
        names = []
        seen = set()
        for i in range(info['count']):
            p = sect_start + i * info['rec_size'] + info['name_offset']
            nb = raw[p:p+64]
            null = nb.find(0)
            name = nb[:null if null >= 0 else 64].decode('ascii', errors='replace').strip()
            if not name or name in seen:
                continue
            seen.add(name)
            names.append(name)
        return names

    input_names  = read_names('InputChannel')
    mix_names    = read_names('Mix')
    matrix_names = read_names('Matrix')
    stereo_names = read_names('Stereo')

    # ── 5. Build the standard result structure ─────────────────────────────
    result = {'inputs': [], 'aux': [], 'groups': [], 'matrix': []}

    for i, name in enumerate(input_names):
        result['inputs'].append({'number': str(i + 1), 'name': name, 'type': 'inputs'})

    for i, name in enumerate(mix_names):
        result['aux'].append({'number': f'MX{i+1}', 'name': name, 'type': 'aux'})

    for i, name in enumerate(stereo_names):
        result['groups'].append({'number': f'ST{i+1}', 'name': name, 'type': 'groups'})

    for i, name in enumerate(matrix_names):
        result['matrix'].append({'number': f'MT{i+1}', 'name': name, 'type': 'matrix'})

    print(f"\n=== RIVAGE PARSE SUMMARY ===")
    print(f"Inputs: {len(result['inputs'])}")
    print(f"Mix (Aux): {len(result['aux'])}")
    print(f"Stereo (Groups): {len(result['groups'])}")
    print(f"Matrix: {len(result['matrix'])}")

    return result


def parse_dlive_show_file(file_content):
    """Parse an Allen & Heath dLive .tar.gz show file and extract channel sections.

    The show file is a tar.gz archive containing a Show/ directory.
    Channel names live in Show/Scenes/StageBoxScene65535.tar.gz, which itself
    contains a .dat binary file with "Name Colour Manager" sections.
    Each channel record is 9 bytes: 1 byte color index + 8 bytes null-padded name.
    """

    if isinstance(file_content, str):
        file_content = file_content.encode('latin-1')

    # ── 1. Open outer .tar.gz and extract StageBoxScene65535.tar.gz ───────
    try:
        outer = tarfile.open(fileobj=io.BytesIO(file_content), mode='r:gz')
    except Exception:
        return {'inputs': [], 'aux': [], 'groups': [], 'matrix': []}

    scene_gz = None
    for member in outer.getmembers():
        if 'StageBoxScene65535' in member.name and member.name.endswith('.tar.gz'):
            f = outer.extractfile(member)
            if f:
                scene_gz = f.read()
            break
    outer.close()

    if scene_gz is None:
        return {'inputs': [], 'aux': [], 'groups': [], 'matrix': []}

    # ── 2. Open inner .tar.gz and extract the .dat file ───────────────────
    try:
        inner = tarfile.open(fileobj=io.BytesIO(scene_gz), mode='r:gz')
    except Exception:
        return {'inputs': [], 'aux': [], 'groups': [], 'matrix': []}

    dat = None
    for member in inner.getmembers():
        if member.name.endswith('.dat'):
            f = inner.extractfile(member)
            if f:
                dat = f.read()
            break
    inner.close()

    if dat is None:
        return {'inputs': [], 'aux': [], 'groups': [], 'matrix': []}

    # ── 3. Find all Name Colour Manager sections and parse records ─────────
    # ALL_SECTIONS is used purely for boundary detection — every section header
    # in the file must be listed so gaps between wanted sections are accurate.
    # SECTION_MAP controls which sections we actually extract.
    ALL_SECTIONS = [
        b'#Input Channel Name Colour Manager',
        b'Mono Group Channel Name Colour Manager',
        b'Stereo Group Channel Name Colour Manager',
        b'Mono Aux Channel Name Colour Manager',
        b'Stereo Aux Channel Name Colour Manager',
        b'Mono FX Send Channel Name Colour Manager',
        b'Stereo FX Send Channel Name Colour Manager',
        b'Stereo AHFX Send Channel Name Colour Manager',
        b'Main Channel Name Colour Manager',
        b'Mono Matrix Channel Name Colour Manager',
        b'Stereo Matrix Channel Name Colour Manager',
        b'FX Return Channel Name Colour Manager',
        b'AHFX Return Channel Name Colour Manager',
        b'DCA Channel Name Colour Manager',
        b'Monitor Channel Name Colour Manager',
    ]
    SECTION_MAP = {
        b'#Input Channel Name Colour Manager':        'inputs',
        b'Mono Group Channel Name Colour Manager':    'groups',
        b'Stereo Group Channel Name Colour Manager':  'groups',
        b'Mono Aux Channel Name Colour Manager':      'aux',
        b'Stereo Aux Channel Name Colour Manager':    'aux',
        b'Main Channel Name Colour Manager':          'groups',
        b'Mono Matrix Channel Name Colour Manager':   'matrix',
        b'Stereo Matrix Channel Name Colour Manager': 'matrix',
        b'Monitor Channel Name Colour Manager':       'aux',
        b'DCA Channel Name Colour Manager':           'groups',
    }

    # Build sorted list of (name_pos, data_start, section_name) for every section found
    all_found = []
    for section_name in ALL_SECTIONS:
        pos = dat.find(section_name + b'\x00')
        if pos >= 0:
            data_start = pos + len(section_name) + 1
            all_found.append((pos, data_start, section_name))
    all_found.sort()

    STEREO_SECTIONS = {
        b'Stereo Group Channel Name Colour Manager',
        b'Stereo Aux Channel Name Colour Manager',
        b'Stereo Matrix Channel Name Colour Manager',
    }

    # Label used when a channel has only a numeric default name (e.g. "1", "32")
    DEFAULT_LABEL = {
        b'#Input Channel Name Colour Manager':        'Input',
        b'Mono Group Channel Name Colour Manager':    'Mono Grp',
        b'Stereo Group Channel Name Colour Manager':  'Stereo Grp',
        b'Mono Aux Channel Name Colour Manager':      'Mono Aux',
        b'Stereo Aux Channel Name Colour Manager':    'Stereo Aux',
        b'Main Channel Name Colour Manager':          'Main',
        b'Mono Matrix Channel Name Colour Manager':   'Mono Mtx',
        b'Stereo Matrix Channel Name Colour Manager': 'Stereo Mtx',
        b'Monitor Channel Name Colour Manager':       'Monitor',
        b'DCA Channel Name Colour Manager':           'DCA',
    }

    result = {'inputs': [], 'aux': [], 'groups': [], 'matrix': []}
    counters = {'inputs': 0, 'aux': 0, 'groups': 0, 'matrix': 0}
    prefix_map = {'inputs': '', 'aux': 'AUX', 'groups': 'GRP', 'matrix': 'MTX'}

    for idx, (pos, data_start, section_name) in enumerate(all_found):
        category = SECTION_MAP.get(section_name)
        if category is None:
            continue  # boundary only — don't extract

        # Use the very next section (wanted or not) as the boundary
        next_pos = all_found[idx + 1][0] if idx + 1 < len(all_found) else data_start + 9 * 256
        count = min((next_pos - data_start) // 9, 256)
        default_label = DEFAULT_LABEL.get(section_name, '')
        section_idx = 0

        for i in range(count):
            rec = dat[data_start + i * 9: data_start + i * 9 + 9]
            if len(rec) < 9:
                break
            name_bytes = rec[1:9]
            null = name_bytes.find(0)
            raw_name = name_bytes[:null if null >= 0 else 8]
            try:
                name = raw_name.decode('ascii').strip()
            except UnicodeDecodeError:
                break  # non-ASCII = past end of section
            if not name or not name.isprintable():
                break  # padding or control bytes = past end of section

            section_idx += 1
            # Replace bare numeric default names with a descriptive label
            if name.isdigit():
                name = f'{default_label} {section_idx}'

            counters[category] += 1
            n = counters[category]
            is_stereo = section_name in STEREO_SECTIONS
            if category == 'inputs':
                number = str(n)
            elif is_stereo:
                number = f'{prefix_map[category]}{n}s'
            else:
                number = f'{prefix_map[category]}{n}'
            result[category].append({'number': number, 'name': name, 'type': category})

    print(f"\n=== DLIVE PARSE SUMMARY ===")
    print(f"Inputs: {len(result['inputs'])}")
    print(f"Aux: {len(result['aux'])}")
    print(f"Groups: {len(result['groups'])}")
    print(f"Matrix: {len(result['matrix'])}")

    return result


def hex_to_reaper_color(hex_color):
    """Convert #RRGGBB to REAPER PEAKCOL integer (R|G<<8|B<<16)|0x1000000"""
    hex_color = hex_color.lstrip('#')
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return (r | (g << 8) | (b << 16)) | 0x1000000


def _track_block(name, peakcol, rec_input, nchan):
    """Build a single REAPER TRACK block."""
    track_id = str(uuid.uuid4()).upper()
    return [
        '<TRACK',
        f'  NAME "{name}"',
        f'  PEAKCOL {peakcol}',
        '  BEAT -1',
        '  AUTOMODE 0',
        '  VOLPAN 1 0 -1 -1 1',
        '  MUTESOLO 0 0 0',
        '  IPHASE 0',
        '  PLAYOFFS 0 1',
        '  ISBUS 0 0',
        '  BUSCOMP 0 0 0 0 0',
        '  SHOWINMIX 1 0.6667 0.5 1 0.5 0 0 0',
        f'  REC 1 {rec_input} 1 0 0 0 0 0',
        '  VU 2',
        '  TRACKHEIGHT 0 0 0 0 0 0',
        '  INQ 0 0 0 0.5 100 0 0 100',
        f'  NCHAN {nchan}',
        '  FX 1',
        f'  TRACKID {{{track_id}}}',
        '  PERF 0',
        '  MIDIOUT -1',
        '  MAINSEND 1 0',
        '>',
    ]


def generate_reaper_track_template(channels, stereo_mode='split'):
    """Generate Reaper track template from channel list with sequential input routing."""

    template_lines = []
    hw = 0  # 0-based hardware input counter

    for channel in channels:
        name    = channel['name']
        peakcol = hex_to_reaper_color(channel['color']) if channel.get('color') else 16576
        is_stereo = channel['number'].endswith('s')

        if is_stereo and stereo_mode == 'stereo':
            # One stereo track — input encoded as 1024 + left_input_index
            template_lines.extend(_track_block(name, peakcol, 1024 + hw, 2))
            hw += 2
        elif is_stereo:
            # Two mono tracks (L then R), each consuming one input
            for suffix in [' L', ' R']:
                template_lines.extend(_track_block(name + suffix, peakcol, hw, 1))
                hw += 1
        else:
            # Single mono track
            template_lines.extend(_track_block(name, peakcol, hw, 1))
            hw += 1

    return '\n'.join(template_lines)


def find_available_port(start_port=8081, max_attempts=10):
    """Find an available port starting from start_port"""
    for port in range(start_port, start_port + max_attempts):
        try:
            # Try to bind to the port
            test_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            test_socket.bind(('localhost', port))
            test_socket.close()
            return port
        except OSError:
            continue
    return None


class DiGiCoToReaperHandler(BaseHTTPRequestHandler):
    
    def log_message(self, format, *args):
        """Suppress request logging"""
        pass
    
    def do_GET(self):
        """Serve the web interface"""
        if self.path == '/':
            self.serve_html()
        elif self.path == '/heartbeat':
            # Simple heartbeat endpoint
            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.end_headers()
            self.wfile.write(b'{"status":"ok"}')
        else:
            self.send_error(404)
    
    def do_POST(self):
        """Handle file upload and conversion"""
        if self.path == '/convert':
            self.handle_conversion()
        elif self.path == '/generate':
            self.handle_generate()
        else:
            self.send_error(404)
    
    def serve_html(self):
        """Serve the main HTML interface"""
        html = '''
<!DOCTYPE html>
<html>
<head>
    <title>DiGiCo to Reaper Converter</title>
    <meta charset="UTF-8">
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif;
            background: #f5f5f7;
            padding: 40px 20px;
            transition: background 0.3s;
        }
        body.dark {
            background: #000;
        }
        body.dark .container {
            background: #1a1a1a;
            box-shadow: 0 2px 20px rgba(0,0,0,0.6);
        }
        body.dark h1, body.dark .track-name, body.dark .track-number,
        body.dark label, body.dark span, body.dark p, body.dark li {
            color: #e0e0e0;
        }
        body.dark .subtitle, body.dark .credit { color: #888; }
        body.dark .preview-list, body.dark .info-box,
        body.dark #sectionSelector { background: #111; }
        body.dark .track-item { background: #222; border-color: #333; }
        body.dark .track-item:hover { background: #2a2a2a; }
        body.dark .track-item.active-highlight { background: #1a2a3a; }
        body.dark .upload-area { background: #111; border-color: #444; }
        body.dark .upload-area:hover { background: #1a1a1a; border-color: #007aff; }
        body.dark .upload-text { color: #ccc; }
        body.dark .upload-subtext { color: #666; }
        body.dark .stereo-toggle-options { background: #2a2a2a; }
        body.dark .stereo-option { color: #888; }
        body.dark .stereo-option.active { background: #444; color: #fff; }
        .dark-mode-btn {
            position: fixed;
            top: 16px;
            right: 20px;
            background: #1d1d1f;
            color: white;
            border: none;
            border-radius: 20px;
            padding: 8px 16px;
            font-size: 13px;
            font-weight: 500;
            cursor: pointer;
            z-index: 999;
            transition: background 0.2s;
        }
        .dark-mode-btn:hover { background: #333; }
        body.dark .dark-mode-btn { background: #444; }
        body.dark .dark-mode-btn:hover { background: #555; }
        .container {
            max-width: 1000px;
            margin: 0 auto;
            background: white;
            border-radius: 12px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            padding: 40px;
            transition: background 0.3s, box-shadow 0.3s;
        }
        h1 {
            color: #1d1d1f;
            margin-bottom: 10px;
            font-size: 32px;
            font-weight: 600;
        }
        .subtitle {
            color: #86868b;
            margin-bottom: 40px;
            font-size: 16px;
        }
        .credit {
            font-size: 12px;
            color: #86868b;
            margin-bottom: 15px;
            line-height: 1.5;
        }
        .credit a {
            color: #007aff;
            text-decoration: none;
        }
        .credit a:hover {
            text-decoration: underline;
        }
        .tab-bar {
            display: flex;
            align-items: center;
            gap: 4px;
            margin-bottom: 24px;
            border-bottom: 2px solid #e5e5ea;
            padding-bottom: 0;
            flex-wrap: wrap;
        }
        .tab {
            padding: 8px 16px;
            border-radius: 8px 8px 0 0;
            border: 1px solid transparent;
            border-bottom: none;
            cursor: pointer;
            font-size: 14px;
            font-weight: 500;
            color: #86868b;
            background: none;
            position: relative;
            bottom: -2px;
            transition: background 0.15s, color 0.15s;
            user-select: none;
            white-space: nowrap;
        }
        .tab:hover { background: #f2f2f7; color: #1d1d1f; }
        .tab.active {
            background: white;
            color: #1d1d1f;
            border-color: #e5e5ea;
            border-bottom-color: white;
        }
        body.dark .tab-bar { border-bottom-color: #333; }
        body.dark .tab:hover { background: #2a2a2a; color: #e0e0e0; }
        body.dark .tab.active { background: #1a1a1a; color: #fff; border-color: #333; border-bottom-color: #1a1a1a; }
        .tab-name { outline: none; }
        .tab-close {
            display: inline-block;
            margin-left: 6px;
            color: #c0c0c0;
            font-size: 12px;
            line-height: 1;
            border-radius: 50%;
            width: 14px;
            height: 14px;
            text-align: center;
        }
        .tab-close:hover { background: #ffdddd; color: #c7251a; }
        .tab-add {
            padding: 6px 12px;
            border-radius: 8px;
            border: 1px dashed #c0c0c0;
            background: none;
            color: #86868b;
            cursor: pointer;
            font-size: 18px;
            line-height: 1;
            transition: background 0.15s, color 0.15s;
        }
        .tab-add:hover { background: #f2f2f7; color: #007aff; border-color: #007aff; }
        body.dark .tab-add { border-color: #444; color: #666; }
        body.dark .tab-add:hover { background: #222; color: #007aff; }
        .upload-area {
            border: 3px dashed #d2d2d7;
            border-radius: 12px;
            padding: 60px 40px;
            text-align: center;
            background: #f9f9f9;
            cursor: pointer;
            transition: all 0.3s;
            margin-bottom: 30px;
        }
        .upload-area:hover {
            border-color: #007aff;
            background: #f0f8ff;
        }
        .upload-area.dragover {
            border-color: #007aff;
            background: #e6f2ff;
        }
        .upload-icon {
            font-size: 48px;
            margin-bottom: 20px;
        }
        .upload-text {
            font-size: 18px;
            color: #1d1d1f;
            margin-bottom: 10px;
        }
        .upload-subtext {
            font-size: 14px;
            color: #86868b;
        }
        input[type="file"] {
            display: none;
        }
        .preview-area {
            display: none;
            margin-top: 30px;
        }
        .preview-header {
            font-size: 18px;
            font-weight: 600;
            margin-bottom: 15px;
            color: #1d1d1f;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .select-buttons {
            display: flex;
            gap: 10px;
        }
        .select-buttons button {
            padding: 6px 12px;
            font-size: 13px;
            background: #e5e5ea;
            border: none;
            border-radius: 6px;
            cursor: pointer;
        }
        .select-buttons button:hover {
            background: #d1d1d6;
        }
        .preview-list {
            background: #f9f9f9;
            border-radius: 8px;
            padding: 20px;
            max-height: 500px;
            overflow-y: auto;
            font-family: 'Monaco', 'Menlo', monospace;
            font-size: 13px;
        }
        .track-item {
            padding: 10px;
            border-bottom: 1px solid #e5e5ea;
            display: flex;
            align-items: center;
            gap: 12px;
            cursor: move;
            user-select: none;
        }
        .track-item:last-child {
            border-bottom: none;
        }
        .track-item:hover {
            background: #f0f0f0;
        }
        .track-item.dragging {
            opacity: 0.4;
        }
        .track-item.drop-above {
            border-top: 2px solid #007aff;
        }
        .track-item.drop-below {
            border-bottom: 2px solid #007aff;
        }
        .drag-handle {
            color: #86868b;
            font-size: 18px;
            cursor: grab;
        }
        .drag-handle:active {
            cursor: grabbing;
        }
        .track-checkbox {
            width: 18px;
            height: 18px;
            cursor: pointer;
        }
        .track-number {
            display: inline-block;
            width: 50px;
            color: #86868b;
            font-weight: 600;
        }
        .track-badge {
            display: inline-block;
            padding: 2px 8px;
            margin-right: 10px;
            border-radius: 4px;
            font-size: 11px;
            font-weight: 600;
            text-transform: uppercase;
        }
        .badge-inputs {
            background: #e5e5e7;
            color: #1d1d1f;
        }
        .badge-aux {
            background: #d1e7ff;
            color: #0066cc;
        }
        .badge-groups {
            background: #d4edda;
            color: #155724;
        }
        .badge-matrix {
            background: #fff3cd;
            color: #856404;
        }
        .badge-custom {
            background: #ede9fe;
            color: #7c3aed;
        }
        .track-delete {
            color: #86868b;
            background: none;
            border: none;
            cursor: pointer;
            font-size: 18px;
            padding: 0 4px;
            line-height: 1;
            border-radius: 4px;
            transition: color 0.2s, background 0.2s;
        }
        .track-delete:hover {
            color: #c7251a;
            background: #ffd3d0;
        }
        .track-item.active-highlight {
            background: #e8f0fe;
        }
        .track-item.active-highlight:hover {
            background: #dce8fd;
        }
        .bulk-bar {
            display: none;
            align-items: center;
            gap: 10px;
            padding: 10px 15px;
            background: #007aff;
            color: white;
            border-radius: 8px;
            margin-bottom: 10px;
            font-size: 14px;
            font-weight: 500;
            flex-wrap: wrap;
        }
        .bulk-bar.visible {
            display: flex;
        }
        .bulk-bar-btn {
            background: rgba(255,255,255,0.2);
            color: white;
            border: none;
            padding: 5px 12px;
            border-radius: 6px;
            font-size: 13px;
            font-weight: 500;
            cursor: pointer;
            transition: background 0.15s;
        }
        .bulk-bar-btn:hover {
            background: rgba(255,255,255,0.35);
        }
        .bulk-color-swatch {
            width: 26px;
            height: 26px;
            border-radius: 6px;
            border: 2px solid rgba(255,255,255,0.6);
            cursor: pointer;
            background: white;
            position: relative;
            flex-shrink: 0;
            transition: transform 0.15s;
        }
        .bulk-color-swatch:hover {
            transform: scale(1.1);
        }
        .bulk-color-swatch input[type="color"] {
            position: absolute;
            width: 0;
            height: 0;
            opacity: 0;
            pointer-events: none;
        }
        .section-color-swatch {
            width: 18px;
            height: 18px;
            border-radius: 50%;
            border: 2px dashed #c0c0c0;
            background: transparent;
            cursor: pointer;
            position: relative;
            flex-shrink: 0;
            transition: transform 0.15s, border-color 0.2s;
        }
        .section-color-swatch:hover {
            transform: scale(1.2);
            border-color: #007aff;
        }
        .section-color-swatch.has-color {
            border: 2px solid rgba(0,0,0,0.15);
        }
        .section-color-swatch .color-clear-badge {
            display: none;
            position: absolute;
            top: -5px;
            right: -5px;
            width: 13px;
            height: 13px;
            border-radius: 50%;
            background: #c7251a;
            color: white;
            font-size: 9px;
            line-height: 13px;
            text-align: center;
            cursor: pointer;
            z-index: 10;
            font-weight: bold;
        }
        .section-color-swatch.has-color:hover .color-clear-badge {
            display: block;
        }
        .section-color-swatch input[type="color"] {
            position: absolute;
            width: 0;
            height: 0;
            opacity: 0;
            pointer-events: none;
        }
        .stereo-toggle {
            display: flex;
            align-items: center;
            gap: 12px;
            margin-bottom: 12px;
            font-size: 14px;
            color: #1d1d1f;
        }
        .stereo-toggle-label {
            font-weight: 500;
        }
        .stereo-toggle-options {
            display: flex;
            background: #f2f2f7;
            border-radius: 8px;
            padding: 3px;
        }
        .stereo-option {
            padding: 5px 14px;
            border-radius: 6px;
            cursor: pointer;
            font-size: 13px;
            font-weight: 500;
            color: #86868b;
            transition: background 0.15s, color 0.15s;
            user-select: none;
        }
        .stereo-option.active {
            background: white;
            color: #1d1d1f;
            box-shadow: 0 1px 3px rgba(0,0,0,0.12);
        }
        .track-color-btn {
            width: 20px;
            height: 20px;
            border-radius: 50%;
            border: 2px dashed #c0c0c0;
            cursor: pointer;
            flex-shrink: 0;
            position: relative;
            transition: transform 0.15s, border-color 0.2s;
            background: transparent;
        }
        .track-color-btn:hover {
            transform: scale(1.2);
            border-color: #007aff;
        }
        .track-color-btn.has-color {
            border: 2px solid rgba(0,0,0,0.15);
        }
        .track-color-btn input[type="color"] {
            position: absolute;
            width: 0;
            height: 0;
            opacity: 0;
            pointer-events: none;
        }
        .color-clear-badge {
            display: none;
            position: absolute;
            top: -5px;
            right: -5px;
            width: 13px;
            height: 13px;
            border-radius: 50%;
            background: #c7251a;
            color: white;
            font-size: 9px;
            line-height: 13px;
            text-align: center;
            cursor: pointer;
            z-index: 10;
            font-weight: bold;
        }
        .track-color-btn.has-color:hover .color-clear-badge {
            display: block;
        }
        .track-edit {
            color: #86868b;
            background: none;
            border: none;
            cursor: pointer;
            font-size: 14px;
            padding: 0 4px;
            line-height: 1;
            border-radius: 4px;
            opacity: 0;
            transition: opacity 0.2s, color 0.2s, background 0.2s;
        }
        .track-item:hover .track-edit {
            opacity: 1;
        }
        .track-edit:hover {
            color: #007aff;
            background: #e6f2ff;
        }
        .track-name {
            color: #1d1d1f;
            flex: 1;
        }
        .button-group {
            margin-top: 30px;
            display: flex;
            gap: 15px;
            justify-content: center;
        }
        button {
            padding: 12px 30px;
            border: none;
            border-radius: 8px;
            font-size: 16px;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.2s;
        }
        .btn-primary {
            background: #007aff;
            color: white;
        }
        .btn-primary:hover {
            background: #0051d5;
        }
        .btn-secondary {
            background: #e5e5ea;
            color: #1d1d1f;
        }
        .btn-secondary:hover {
            background: #d1d1d6;
        }
        .message {
            padding: 15px;
            border-radius: 8px;
            margin-top: 20px;
            display: none;
        }
        .message.success {
            background: #d1f2dd;
            color: #248a3d;
            display: block;
        }
        .message.error {
            background: #ffd3d0;
            color: #c7251a;
            display: block;
        }
        .info-box {
            background: #e6f2ff;
            border-left: 4px solid #007aff;
            padding: 20px;
            border-radius: 8px;
            margin-top: 30px;
        }
        .info-box h3 {
            color: #007aff;
            margin-bottom: 10px;
            font-size: 16px;
        }
        .info-box ol {
            margin-left: 20px;
            color: #1d1d1f;
        }
        .info-box li {
            margin-bottom: 8px;
        }
        .modal {
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.5);
            z-index: 1000;
            justify-content: center;
            align-items: center;
        }
        .modal.active {
            display: flex;
        }
        .modal-content {
            background: white;
            border-radius: 12px;
            padding: 30px;
            max-width: 500px;
            width: 90%;
            box-shadow: 0 4px 20px rgba(0,0,0,0.3);
        }
        .modal-content h2 {
            font-size: 20px;
            margin-bottom: 20px;
            color: #1d1d1f;
        }
        .modal-content input {
            width: 100%;
            padding: 12px;
            border: 1px solid #d2d2d7;
            border-radius: 6px;
            font-size: 16px;
            margin-bottom: 20px;
        }
        .modal-buttons {
            display: flex;
            gap: 10px;
            justify-content: flex-end;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="credit">
            <button class="dark-mode-btn" id="darkModeBtn" onclick="toggleDarkMode()">🌙 Dark Mode</button>

        Built by Michael Leckrone<br>
            <a href="mailto:leckroneaudio@gmail.com">leckroneaudio@gmail.com</a>
        </div>
        
        <div class="tab-bar" id="tabBar"></div>

        <h1>Console to Reaper Converter</h1>
        <p class="subtitle">Convert DiGiCo or Yamaha Rivage show files to Reaper track templates</p>

        <div id="uploadArea" class="upload-area" onclick="document.getElementById('fileInput').click()">
            <div class="upload-icon">📄</div>
            <div class="upload-text">Drop your show file here</div>
            <div class="upload-subtext">or click to browse &nbsp;·&nbsp; DiGiCo (.rtf) &nbsp;·&nbsp; Yamaha Rivage (.RIVAGEPM)</div>
        </div>

        <input type="file" id="fileInput" accept=".rtf,.RIVAGEPM,.rivagepm,.tar.gz" onchange="handleFile(this.files[0])">
        
        <div id="message" class="message"></div>
        
        <div style="text-align: right; margin: -15px 0 20px;">
            <button class="btn-secondary" onclick="openAddChannelModal()" style="font-size: 14px; padding: 8px 16px;">+ Add Channel Manually</button>
        </div>

        <div class="stereo-toggle">
                <span class="stereo-toggle-label">Stereo channels:</span>
                <div class="stereo-toggle-options">
                    <div class="stereo-option active" id="optSplit" onclick="setStereoMode('split')">Split to L/R Mono</div>
                    <div class="stereo-option" id="optStereo" onclick="setStereoMode('stereo')">Keep Stereo</div>
                </div>
            </div>
            <div class="button-group">
                <button class="btn-primary" onclick="downloadTemplate()">⬇️ Download Track Template</button>
                <button class="btn-secondary" onclick="downloadCSV()">📄 Export as CSV</button>
                <button class="btn-secondary" onclick="reset()">🔄 Upload New File</button>
            </div>

        <div id="previewArea" class="preview-area">
            <div class="preview-header">
                <span>Track Preview (<span id="selectedCount">0</span> of <span id="trackCount">0</span> selected)</span>
                <div class="select-buttons">
                    <button onclick="selectAll()">Select All</button>
                    <button onclick="selectNone()">Deselect All</button>
                    <button onclick="removeUnnamed()" title="Deselect channels that still have their default name">Remove Unnamed</button>
                    <button id="undoBtn" onclick="undo()" disabled style="opacity: 0.4;">↩ Undo</button>
                    <button onclick="openAddChannelModal()" style="background: #007aff; color: white;">+ Add Channel</button>
                </div>
            </div>
            
            <!-- Section Selection -->
            <div id="sectionSelector" style="margin: 20px 0; padding: 15px; background: #f9f9f9; border-radius: 8px;">
                <div style="font-weight: 600; margin-bottom: 5px;">Quick Select Sections:</div>
                <div style="font-size: 13px; color: #666; margin-bottom: 10px;">Check to auto-select all channels in a section. Uncheck to manually pick individual channels.</div>
                <div style="display: flex; gap: 20px; flex-wrap: wrap;">
                    <div style="display: flex; align-items: center; gap: 8px;">
                        <label style="display: flex; align-items: center; gap: 8px; cursor: pointer; margin: 0;">
                            <input type="checkbox" id="includeInputs" checked onchange="updateSectionPreview()" style="width: 18px; height: 18px;">
                            <span>Inputs (<span id="inputsCount">0</span>)</span>
                        </label>
                        <div class="section-color-swatch" id="swatchInputs" title="Set color for all Inputs" onclick="document.getElementById('colorInputs').click()">
                            <input type="color" id="colorInputs" onchange="applyColorToSection('inputs', this.value)">
                            <span class="color-clear-badge" onclick="event.stopPropagation(); clearColorFromSection('inputs')" title="Clear color">✕</span>
                        </div>
                    </div>
                    <div style="display: flex; align-items: center; gap: 8px;">
                        <label style="display: flex; align-items: center; gap: 8px; cursor: pointer; margin: 0;">
                            <input type="checkbox" id="includeAux" onchange="updateSectionPreview()" style="width: 18px; height: 18px;">
                            <span>Aux Outputs (<span id="auxCount">0</span>)</span>
                        </label>
                        <div class="section-color-swatch" id="swatchAux" title="Set color for all Aux Outputs" onclick="document.getElementById('colorAux').click()">
                            <input type="color" id="colorAux" onchange="applyColorToSection('aux', this.value)">
                            <span class="color-clear-badge" onclick="event.stopPropagation(); clearColorFromSection('aux')" title="Clear color">✕</span>
                        </div>
                    </div>
                    <div style="display: flex; align-items: center; gap: 8px;">
                        <label style="display: flex; align-items: center; gap: 8px; cursor: pointer; margin: 0;">
                            <input type="checkbox" id="includeGroups" onchange="updateSectionPreview()" style="width: 18px; height: 18px;">
                            <span>Group Outputs (<span id="groupsCount">0</span>)</span>
                        </label>
                        <div class="section-color-swatch" id="swatchGroups" title="Set color for all Group Outputs" onclick="document.getElementById('colorGroups').click()">
                            <input type="color" id="colorGroups" onchange="applyColorToSection('groups', this.value)">
                            <span class="color-clear-badge" onclick="event.stopPropagation(); clearColorFromSection('groups')" title="Clear color">✕</span>
                        </div>
                    </div>
                    <div style="display: flex; align-items: center; gap: 8px;">
                        <label style="display: flex; align-items: center; gap: 8px; cursor: pointer; margin: 0;">
                            <input type="checkbox" id="includeMatrix" onchange="updateSectionPreview()" style="width: 18px; height: 18px;">
                            <span>Matrix Outputs (<span id="matrixCount">0</span>)</span>
                        </label>
                        <div class="section-color-swatch" id="swatchMatrix" title="Set color for all Matrix Outputs" onclick="document.getElementById('colorMatrix').click()">
                            <input type="color" id="colorMatrix" onchange="applyColorToSection('matrix', this.value)">
                            <span class="color-clear-badge" onclick="event.stopPropagation(); clearColorFromSection('matrix')" title="Clear color">✕</span>
                        </div>
                    </div>
                </div>
            </div>
            
            <!-- Bulk action bar -->
            <div id="bulkBar" class="bulk-bar">
                <span id="bulkCount">0 channels selected</span>
                <div class="bulk-color-swatch" id="bulkColorSwatch" title="Set color for all selected channels">
                    <input type="color" id="bulkColorInput" value="#ff6b6b">
                </div>
                <button class="bulk-bar-btn" onclick="clearBulkColors()">Clear Colors</button>
                <button class="bulk-bar-btn" onclick="clearActiveSelection()" style="margin-left: auto;">✕ Deselect All</button>
            </div>

            <div id="previewList" class="preview-list"></div>

            <div class="stereo-toggle">
                <span class="stereo-toggle-label">Stereo channels:</span>
                <div class="stereo-toggle-options">
                    <div class="stereo-option active" id="optSplitBottom" onclick="setStereoMode('split')">Split to L/R Mono</div>
                    <div class="stereo-option" id="optStereoBottom" onclick="setStereoMode('stereo')">Keep Stereo</div>
                </div>
            </div>
            <div class="button-group">
                <button class="btn-primary" onclick="downloadTemplate()">⬇️ Download Track Template</button>
                <button class="btn-secondary" onclick="downloadCSV()">📄 Export as CSV</button>
                <button class="btn-secondary" onclick="reset()">🔄 Upload New File</button>
            </div>
        </div>
        
        <div class="info-box">
            <h3>How to use:</h3>
            <ol>
                <li><strong>DiGiCo:</strong> Export session report from the console (.rtf file)</li>
                <li><strong>Yamaha Rivage PM:</strong> Copy the .RIVAGEPM show file from the console or Rivage PM Editor</li>
                <li><strong>Allen &amp; Heath dLive:</strong> Export the show file from dLive Director (.tar.gz)</li>
                <li>Upload or drag the file here</li>
                <li>Select/deselect channels you want to import</li>
                <li>Download the .RTrackTemplate file</li>
                <li>Open blank Reaper session</li>
                <li>Track → Insert tracks from template → Select the downloaded file (or drag it in)</li>
            </ol>
        </div>
    </div>
    
    <!-- Filename Modal -->
    <div id="filenameModal" class="modal">
        <div class="modal-content">
            <h2>Save Track Template</h2>
            <input type="text" id="filenameInput" placeholder="Enter filename" value="DiGiCo_Tracks">
            <div class="modal-buttons">
                <button class="btn-secondary" onclick="closeFilenameModal()">Cancel</button>
                <button class="btn-primary" onclick="confirmDownload()">Download</button>
            </div>
        </div>
    </div>
    
    <!-- Add Channel Modal -->
    <div id="addChannelModal" class="modal">
        <div class="modal-content">
            <h2>Add Channel</h2>
            <div style="margin-bottom: 15px;">
                <label style="display: block; font-weight: 500; margin-bottom: 6px; color: #1d1d1f;">Channel Name</label>
                <input type="text" id="newChannelName" placeholder="e.g. Kick Drum">
            </div>
            <div style="display: flex; gap: 12px; margin-bottom: 15px;">
                <div style="flex: 1;">
                    <label style="display: block; font-weight: 500; margin-bottom: 6px; color: #1d1d1f;">Type</label>
                    <select id="newChannelType" style="width: 100%; padding: 12px; border: 1px solid #d2d2d7; border-radius: 6px; font-size: 16px; margin-bottom: 0;">
                        <option value="inputs">Input</option>
                        <option value="aux">Aux</option>
                        <option value="groups">Group</option>
                        <option value="matrix">Matrix</option>
                        <option value="custom">Custom</option>
                    </select>
                </div>
                <div style="width: 90px;">
                    <label style="display: block; font-weight: 500; margin-bottom: 6px; color: #1d1d1f;">Quantity</label>
                    <input type="number" id="newChannelQty" value="1" min="1" max="128" style="width: 100%; padding: 12px; border: 1px solid #d2d2d7; border-radius: 6px; font-size: 16px; margin-bottom: 0;">
                </div>
            </div>
            <label style="display: flex; align-items: center; gap: 8px; margin-bottom: 6px; cursor: pointer;">
                <input type="checkbox" id="newChannelStereo" style="width: 18px; height: 18px; margin-bottom: 0;">
                <span>Stereo channel</span>
            </label>
            <div id="newChannelQtyHint" style="font-size: 12px; color: #86868b; margin-bottom: 20px;">e.g. "Mic" × 3 → Mic 1, Mic 2, Mic 3</div>
            <div class="modal-buttons">
                <button class="btn-secondary" onclick="closeAddChannelModal()">Cancel</button>
                <button class="btn-primary" onclick="addCustomChannel()">Add Channel</button>
            </div>
        </div>
    </div>

    <!-- Disconnect Overlay -->
    <div id="disconnectOverlay" style="display: none; position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.85); z-index: 10000; align-items: center; justify-content: center;">
        <div style="background: white; padding: 40px; border-radius: 12px; text-align: center; max-width: 400px;">
            <div style="font-size: 48px; margin-bottom: 20px;">⚠️</div>
            <h2 style="margin-bottom: 10px; color: #1d1d1f;">Server Disconnected</h2>
            <p style="color: #86868b; margin-bottom: 20px;">The DiGiCo to Reaper converter has been closed.</p>
            <p style="color: #86868b; font-size: 14px;">You can close this tab or restart the app to continue.</p>
        </div>
    </div>
    
    <script>
        // ── Tab / Session management ────────────────────────────────────────
        function newSessionState(name) {
            return {
                name,
                parsedSections: { inputs: [], aux: [], groups: [], matrix: [] },
                currentCombinedChannels: [],
                selectedChannels: new Set(),
                customChannelCount: 0,
                undoStack: [],
                activeChannels: new Set(),
                lastActiveIdx: null,
                sectionChecks: { inputs: true, aux: false, groups: false, matrix: false },
            };
        }

        let sessions = [newSessionState('Session 1')];
        let activeTab = 0;

        function getTab() { return sessions[activeTab]; }

        // Proxy globals so all existing code keeps working without changes
        let parsedSections,
            currentCombinedChannels,
            selectedChannels,
            customChannelCount,
            undoStack,
            activeChannels,
            lastActiveIdx;

        function loadTabState() {
            const s = getTab();
            parsedSections          = s.parsedSections;
            currentCombinedChannels = s.currentCombinedChannels;
            selectedChannels        = s.selectedChannels;
            customChannelCount      = s.customChannelCount;
            undoStack               = s.undoStack;
            activeChannels          = s.activeChannels;
            lastActiveIdx           = s.lastActiveIdx;
            // Restore section checkboxes
            document.getElementById('includeInputs').checked = s.sectionChecks.inputs;
            document.getElementById('includeAux').checked    = s.sectionChecks.aux;
            document.getElementById('includeGroups').checked = s.sectionChecks.groups;
            document.getElementById('includeMatrix').checked = s.sectionChecks.matrix;
        }

        function saveTabState() {
            const s = getTab();
            s.parsedSections          = parsedSections;
            s.currentCombinedChannels = currentCombinedChannels;
            s.selectedChannels        = selectedChannels;
            s.customChannelCount      = customChannelCount;
            s.undoStack               = undoStack;
            s.activeChannels          = activeChannels;
            s.lastActiveIdx           = lastActiveIdx;
            s.sectionChecks = {
                inputs:  document.getElementById('includeInputs').checked,
                aux:     document.getElementById('includeAux').checked,
                groups:  document.getElementById('includeGroups').checked,
                matrix:  document.getElementById('includeMatrix').checked,
            };
        }

        function switchTab(idx) {
            saveTabState();
            activeTab = idx;
            loadTabState();
            renderTabs();
            // Restore UI state for this tab
            const s = getTab();
            const hasChannels = s.currentCombinedChannels.length > 0;
            document.getElementById('previewArea').style.display = hasChannels ? 'block' : 'none';
            document.getElementById('message').style.display = 'none';
            document.getElementById('fileInput').value = '';
            if (hasChannels) {
                showPreview(s.currentCombinedChannels);
                refreshSectionCounts();
            }
            updateUndoBtn();
            updateBulkBar();
        }

        function addTab() {
            saveTabState();
            const n = sessions.length + 1;
            sessions.push(newSessionState('Session ' + n));
            activeTab = sessions.length - 1;
            loadTabState();
            renderTabs();
            // Clear the UI for the fresh tab
            document.getElementById('previewArea').style.display = 'none';
            document.getElementById('message').style.display = 'none';
            document.getElementById('fileInput').value = '';
            updateUndoBtn();
            updateBulkBar();
        }

        function closeTab(idx) {
            if (sessions.length === 1) return; // keep at least one
            sessions.splice(idx, 1);
            if (activeTab >= sessions.length) activeTab = sessions.length - 1;
            loadTabState();
            renderTabs();
            const s = getTab();
            const hasChannels = s.currentCombinedChannels.length > 0;
            document.getElementById('previewArea').style.display = hasChannels ? 'block' : 'none';
            if (hasChannels) { showPreview(s.currentCombinedChannels); refreshSectionCounts(); }
            updateUndoBtn();
            updateBulkBar();
        }

        function renderTabs() {
            const bar = document.getElementById('tabBar');
            bar.innerHTML = '';
            sessions.forEach((s, i) => {
                const tab = document.createElement('div');
                tab.className = 'tab' + (i === activeTab ? ' active' : '');

                function startTabRename(tabEl) {
                    const existingInput = tabEl.querySelector('input');
                    if (existingInput) return;
                    const span = tabEl.querySelector('.tab-name');
                    if (!span) return;
                    const input = document.createElement('input');
                    input.value = sessions[i].name;
                    input.style.cssText = 'width:' + Math.max(60, sessions[i].name.length * 9) + 'px;font:inherit;border:none;outline:1px solid #007aff;border-radius:3px;padding:0 3px;background:transparent;color:inherit;';
                    tabEl.replaceChild(input, span);
                    input.focus();
                    input.select();
                    let committed = false;
                    function commit() {
                        if (committed) return;
                        committed = true;
                        const val = input.value.trim();
                        sessions[i].name = val || sessions[i].name;
                        tabEl.replaceChild(span, input);
                        span.textContent = sessions[i].name;
                    }
                    input.addEventListener('blur', commit);
                    input.addEventListener('keydown', (ev) => {
                        if (ev.key === 'Enter') { ev.preventDefault(); input.blur(); }
                        if (ev.key === 'Escape') { committed = true; tabEl.replaceChild(span, input); span.textContent = sessions[i].name; }
                        ev.stopPropagation();
                    });
                }

                const nameSpan = document.createElement('span');
                nameSpan.className = 'tab-name';
                nameSpan.textContent = s.name;
                nameSpan.title = 'Double-click to rename';

                let clickTimer = null;
                tab.addEventListener('click', (e) => {
                    if (e.target.tagName === 'INPUT' || e.target.classList.contains('tab-close')) return;
                    if (clickTimer) {
                        // Double-click detected
                        clearTimeout(clickTimer);
                        clickTimer = null;
                        if (i !== activeTab) { switchTab(i); setTimeout(() => startTabRename(document.querySelectorAll('.tab')[i]), 50); }
                        else startTabRename(tab);
                    } else {
                        clickTimer = setTimeout(() => { clickTimer = null; if (i !== activeTab) switchTab(i); }, 220);
                    }
                });

                tab.appendChild(nameSpan);

                if (sessions.length > 1) {
                    const close = document.createElement('span');
                    close.className = 'tab-close';
                    close.textContent = '×';
                    close.title = 'Close session';
                    close.onclick = (e) => { e.stopPropagation(); closeTab(i); };
                    tab.appendChild(close);
                }

                bar.appendChild(tab);
            });

            const addBtn = document.createElement('button');
            addBtn.className = 'tab-add';
            addBtn.textContent = '+';
            addBtn.title = 'New session';
            addBtn.onclick = addTab;
            bar.appendChild(addBtn);
        }

        // Initialise
        loadTabState();
        renderTabs();

        // ── Non-session globals ──────────────────────────────────────────────
        let draggedIndex = null;
        let pendingDownloadType = 'template';

        function saveUndo() {
            undoStack.push({
                channels: currentCombinedChannels.map(ch => ({ ...ch })),
                selected: new Set(selectedChannels),
                customCount: customChannelCount
            });
            if (undoStack.length > 50) undoStack.shift();
            // updateUndoBtn may not exist yet on first call — defer safely
            setTimeout(updateUndoBtn, 0);
        }

        function undo() {
            if (undoStack.length === 0) return;
            const prev = undoStack.pop();
            currentCombinedChannels.length = 0;
            currentCombinedChannels.push(...prev.channels);
            selectedChannels = prev.selected;
            customChannelCount = prev.customCount;
            activeChannels.clear();
            showPreview(currentCombinedChannels);
            refreshSectionCounts();
            updateBulkBar();
            updateUndoBtn();
        }

        function updateUndoBtn() {
            const btn = document.getElementById('undoBtn');
            if (!btn) return;
            btn.disabled = undoStack.length === 0;
            btn.style.opacity = undoStack.length === 0 ? '0.4' : '1';
        }
        let stereoMode = 'split';
        let didDrag = false;
        
        // Drag and drop
        const uploadArea = document.getElementById('uploadArea');
        
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        });
        
        uploadArea.addEventListener('dragleave', () => {
            uploadArea.classList.remove('dragover');
        });
        
        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
            const file = e.dataTransfer.files[0];
            const name = file ? file.name.toLowerCase() : '';
            if (file && (name.endsWith('.rtf') || name.endsWith('.rivagepm') || name.endsWith('.tar.gz'))) {
                handleFile(file);
            } else {
                showMessage('Please upload a .rtf (DiGiCo) or .RIVAGEPM (Yamaha Rivage) file', 'error');
            }
        });
        
        function handleFile(file) {
            if (!file) return;
            
            const formData = new FormData();
            formData.append('file', file);
            
            showMessage('Processing file...', 'success');
            
            fetch('/convert', {
                method: 'POST',
                body: formData
            })
            .then(r => r.json())
            .then(data => {
                if (data.success) {
                    // Store all sections
                    parsedSections = data.sections;
                    
                    // Update section counts
                    document.getElementById('inputsCount').textContent = data.counts.inputs;
                    document.getElementById('auxCount').textContent = data.counts.aux;
                    document.getElementById('groupsCount').textContent = data.counts.groups;
                    document.getElementById('matrixCount').textContent = data.counts.matrix;
                    
                    // Show/hide checkboxes based on what's available
                    document.getElementById('includeInputs').disabled = data.counts.inputs === 0;
                    document.getElementById('includeAux').disabled = data.counts.aux === 0;
                    document.getElementById('includeGroups').disabled = data.counts.groups === 0;
                    document.getElementById('includeMatrix').disabled = data.counts.matrix === 0;
                    
                    // Update preview with selected sections
                    updateSectionPreview();

                    const total = data.counts.inputs + data.counts.aux + data.counts.groups + data.counts.matrix;
                    if (total === 0) {
                        showMessage('✗ "' + file.name + '" — No Channels Found, Please ensure Include: Channels is selected when saving the Session Report.', 'error');
                    } else {
                        showMessage(`✓ "${file.name}" — ${total} total channels (${data.counts.inputs} inputs, ${data.counts.aux} aux, ${data.counts.groups} groups, ${data.counts.matrix} matrix)`, 'success');
                    }
                } else {
                    showMessage('✗ Error: ' + data.error, 'error');
                }
            })
            .catch(err => {
                showMessage('✗ Error processing file: ' + err, 'error');
            });
        }
        
        function updateSectionPreview() {
            // ALWAYS show all channels from all sections
            let allChannels = [].concat(
                parsedSections.inputs,
                parsedSections.aux,
                parsedSections.groups,
                parsedSections.matrix
            );
            
            // Store combined channels
            currentCombinedChannels = allChannels;
            
            // Auto-select only channels from CHECKED sections
            selectedChannels.clear();
            
            let currentIndex = 0;
            
            // Select inputs if checked
            if (document.getElementById('includeInputs').checked) {
                for (let i = 0; i < parsedSections.inputs.length; i++) {
                    selectedChannels.add(currentIndex + i);
                }
            }
            currentIndex += parsedSections.inputs.length;
            
            // Select aux if checked
            if (document.getElementById('includeAux').checked) {
                for (let i = 0; i < parsedSections.aux.length; i++) {
                    selectedChannels.add(currentIndex + i);
                }
            }
            currentIndex += parsedSections.aux.length;
            
            // Select groups if checked
            if (document.getElementById('includeGroups').checked) {
                for (let i = 0; i < parsedSections.groups.length; i++) {
                    selectedChannels.add(currentIndex + i);
                }
            }
            currentIndex += parsedSections.groups.length;
            
            // Select matrix if checked
            if (document.getElementById('includeMatrix').checked) {
                for (let i = 0; i < parsedSections.matrix.length; i++) {
                    selectedChannels.add(currentIndex + i);
                }
            }
            
            showPreview(allChannels);
        }
        
        function showPreview(channels) {
            const previewArea = document.getElementById('previewArea');
            const previewList = document.getElementById('previewList');
            const trackCount = document.getElementById('trackCount');
            
            previewList.innerHTML = '';
            
            channels.forEach((ch, idx) => {
                const div = document.createElement('div');
                div.className = 'track-item' + (activeChannels.has(idx) ? ' active-highlight' : '');
                div.draggable = true;
                div.dataset.index = idx;

                // Drag events
                div.addEventListener('dragstart', handleDragStart);
                div.addEventListener('dragover', handleDragOver);
                div.addEventListener('drop', handleDrop);
                div.addEventListener('dragend', handleDragEnd);

                // Row click for multi-select (ignore interactive children)
                div.addEventListener('click', (e) => {
                    if (didDrag) return;
                    if (e.target.closest('input, button, .track-color-btn, .drag-handle')) return;
                    handleRowClick(e, idx);
                });
                
                // Drag handle
                const dragHandle = document.createElement('span');
                dragHandle.className = 'drag-handle';
                dragHandle.textContent = '☰';
                
                const checkbox = document.createElement('input');
                checkbox.type = 'checkbox';
                checkbox.className = 'track-checkbox';
                checkbox.dataset.idx = idx;
                checkbox.checked = selectedChannels.has(idx);
                checkbox.onchange = () => toggleChannel(idx);
                
                const number = document.createElement('span');
                number.className = 'track-number';
                number.textContent = ch.number;
                
                // Add type badge
                const badge = document.createElement('span');
                badge.className = 'track-badge';
                
                // Determine badge text and color based on type
                const badgeInfo = {
                    'inputs': { text: 'IN', class: 'badge-inputs' },
                    'aux': { text: 'AUX', class: 'badge-aux' },
                    'groups': { text: 'GRP', class: 'badge-groups' },
                    'matrix': { text: 'MTX', class: 'badge-matrix' },
                    'custom': { text: 'CUST', class: 'badge-custom' }
                };
                
                const info = badgeInfo[ch.type] || { text: 'CH', class: 'badge-inputs' };
                badge.textContent = info.text;
                badge.classList.add(info.class);
                
                const name = document.createElement('span');
                name.className = 'track-name';
                name.textContent = ch.name;
                
                name.addEventListener('dblclick', (e) => { e.stopPropagation(); startEditName(idx, name, ch); });
                name.title = 'Double-click to rename';

                // Color swatch
                const colorBtn = document.createElement('div');
                colorBtn.className = 'track-color-btn' + (ch.color ? ' has-color' : '');
                if (ch.color) colorBtn.style.backgroundColor = ch.color;
                colorBtn.title = 'Click to set track color';

                const colorInput = document.createElement('input');
                colorInput.type = 'color';
                colorInput.value = ch.color || '#ff6b6b';
                colorBtn.appendChild(colorInput);

                // Clear badge — appears on hover when color is set
                const clearBadge = document.createElement('span');
                clearBadge.className = 'color-clear-badge';
                clearBadge.textContent = '✕';
                clearBadge.title = 'Clear color';
                colorBtn.appendChild(clearBadge);

                colorBtn.addEventListener('click', (e) => { e.stopPropagation(); colorInput.click(); });

                clearBadge.addEventListener('click', (e) => {
                    e.stopPropagation();
                    currentCombinedChannels[idx].color = null;
                    colorBtn.style.backgroundColor = '';
                    colorBtn.classList.remove('has-color');
                    colorInput.value = '#ff6b6b';
                });

                colorInput.addEventListener('change', (e) => {
                    currentCombinedChannels[idx].color = e.target.value;
                    colorBtn.style.backgroundColor = e.target.value;
                    colorBtn.classList.add('has-color');
                });

                const editBtn = document.createElement('button');
                editBtn.className = 'track-edit';
                editBtn.textContent = '✎';
                editBtn.title = 'Rename channel';
                editBtn.onclick = (e) => { e.stopPropagation(); startEditName(idx, name, ch); };

                div.appendChild(dragHandle);
                div.appendChild(checkbox);
                div.appendChild(number);
                div.appendChild(badge);
                div.appendChild(colorBtn);
                div.appendChild(name);
                div.appendChild(editBtn);

                const deleteBtn = document.createElement('button');
                deleteBtn.className = 'track-delete';
                deleteBtn.textContent = '×';
                deleteBtn.title = 'Remove channel';
                deleteBtn.onclick = (e) => { e.stopPropagation(); saveUndo(); deleteChannel(idx); };
                div.appendChild(deleteBtn);

                previewList.appendChild(div);
            });
            
            trackCount.textContent = channels.length;
            updateSelectedCount();
            previewArea.style.display = 'block';

            // Click on list background clears active selection
            previewList.addEventListener('click', (e) => {
                if (!e.target.closest('.track-item')) {
                    clearActiveSelection();
                }
            }, { once: true });

            // Wire up bulk color swatch after render
            const bulkSwatch = document.getElementById('bulkColorSwatch');
            const bulkInput = document.getElementById('bulkColorInput');
            bulkSwatch.onclick = () => bulkInput.click();
            bulkInput.onchange = (e) => {
                activeChannels.forEach(i => { currentCombinedChannels[i].color = e.target.value; });
                showPreview(currentCombinedChannels);
                updateBulkBar();
            };
        }
        
        let scrollInterval = null;
        let dropInsertBefore = true; // whether to insert before or after target

        function clearDropIndicators() {
            document.querySelectorAll('.drop-above, .drop-below').forEach(el => {
                el.classList.remove('drop-above', 'drop-below');
            });
        }

        function handleDragStart(e) {
            didDrag = true;
            draggedIndex = parseInt(e.currentTarget.dataset.index);
            if (!activeChannels.has(draggedIndex)) {
                activeChannels.clear();
            }
            e.currentTarget.classList.add('dragging');
            e.dataTransfer.effectAllowed = 'move';
        }

        function handleDragOver(e) {
            e.preventDefault();
            e.dataTransfer.dropEffect = 'move';

            const target = e.target.closest('.track-item');
            clearDropIndicators();
            if (target && target.dataset.index !== undefined) {
                const rect = target.getBoundingClientRect();
                const midY = rect.top + rect.height / 2;
                dropInsertBefore = e.clientY < midY;
                target.classList.add(dropInsertBefore ? 'drop-above' : 'drop-below');
            }

            // Auto-scroll the preview list
            const list = document.getElementById('previewList');
            const listRect = list.getBoundingClientRect();
            const scrollZone = 50;
            if (scrollInterval) { clearInterval(scrollInterval); scrollInterval = null; }
            if (e.clientY < listRect.top + scrollZone) {
                scrollInterval = setInterval(() => { list.scrollTop -= 8; }, 16);
            } else if (e.clientY > listRect.bottom - scrollZone) {
                scrollInterval = setInterval(() => { list.scrollTop += 8; }, 16);
            }

            return false;
        }

        function handleDrop(e) {
            if (e.stopPropagation) e.stopPropagation();
            clearDropIndicators();
            if (scrollInterval) { clearInterval(scrollInterval); scrollInterval = null; }

            const target = e.target.closest('.track-item');
            if (!target || target.dataset.index === undefined) return false;

            let dropIndex = parseInt(target.dataset.index);
            // Adjust insert position based on whether we're dropping above/below
            if (!dropInsertBefore) dropIndex = Math.min(dropIndex + 1, currentCombinedChannels.length - 1);
            saveUndo();
            const dropItem = currentCombinedChannels[dropIndex];

            // Snapshot selection by object reference before reorder
            const selectedObjects = new Set(Array.from(selectedChannels).map(i => currentCombinedChannels[i]));
            const activeObjects = new Set(Array.from(activeChannels).map(i => currentCombinedChannels[i]));

            const isMulti = activeChannels.size > 1 && activeChannels.has(draggedIndex);

            if (isMulti) {
                // Multi-drag: move all active channels together
                const dragging = Array.from(activeObjects);
                const draggingSet = activeObjects;

                if (!draggingSet.has(dropItem)) {
                    const kept = currentCombinedChannels.filter(ch => !draggingSet.has(ch));
                    const insertAt = kept.indexOf(dropItem);
                    currentCombinedChannels.length = 0;
                    if (insertAt === -1) {
                        currentCombinedChannels.push(...kept, ...dragging);
                    } else {
                        currentCombinedChannels.push(...kept.slice(0, insertAt), ...dragging, ...kept.slice(insertAt));
                    }
                }
            } else if (draggedIndex !== dropIndex) {
                // Single drag
                const draggedItem = currentCombinedChannels[draggedIndex];
                currentCombinedChannels.splice(draggedIndex, 1);
                currentCombinedChannels.splice(dropIndex, 0, draggedItem);
            }

            // Rebuild index sets from object references
            selectedChannels.clear();
            activeChannels.clear();
            currentCombinedChannels.forEach((ch, i) => {
                if (selectedObjects.has(ch)) selectedChannels.add(i);
                if (activeObjects.has(ch)) activeChannels.add(i);
            });

            showPreview(currentCombinedChannels);
            updateBulkBar();
            return false;
        }
        
        function handleDragEnd(e) {
            e.currentTarget.classList.remove('dragging');
            clearDropIndicators();
            if (scrollInterval) { clearInterval(scrollInterval); scrollInterval = null; }
            draggedIndex = null;
            setTimeout(() => { didDrag = false; }, 0);
        }
        
        function toggleChannel(idx) {
            if (selectedChannels.has(idx)) {
                selectedChannels.delete(idx);
            } else {
                selectedChannels.add(idx);
            }
            updateSelectedCount();
        }
        
        function updateSelectedCount() {
            document.getElementById('selectedCount').textContent = selectedChannels.size;
        }
        
        function selectAll() {
            selectedChannels.clear();
            currentCombinedChannels.forEach((ch, idx) => {
                selectedChannels.add(idx);
            });

            document.querySelectorAll('.track-checkbox').forEach(cb => { cb.checked = true; });

            // Sync section checkboxes
            ['includeInputs', 'includeAux', 'includeGroups', 'includeMatrix'].forEach(id => {
                document.getElementById(id).checked = true;
            });

            updateSelectedCount();
        }

        function selectNone() {
            selectedChannels.clear();

            document.querySelectorAll('.track-checkbox').forEach(cb => { cb.checked = false; });

            // Sync section checkboxes
            ['includeInputs', 'includeAux', 'includeGroups', 'includeMatrix'].forEach(id => {
                document.getElementById(id).checked = false;
            });

            updateSelectedCount();
        }

        function removeUnnamed() {
            // Deselect channels that still have a default/auto-generated name.
            // Matches: bare numbers ("1", "64"), and our dLive descriptive defaults
            // ("Input 4", "Mono Aux 3", "Stereo Grp 1", "DCA 12", etc.)
            const defaultPattern = /^(Input|Mono Grp|Stereo Grp|Mono Aux|Stereo Aux|Main|Mono Mtx|Stereo Mtx|Monitor|DCA) \d+$|^\d+$/;

            currentCombinedChannels.forEach((ch, idx) => {
                if (defaultPattern.test(ch.name.trim())) {
                    selectedChannels.delete(idx);
                }
            });

            document.querySelectorAll('.track-checkbox').forEach(cb => {
                const idx = parseInt(cb.dataset.idx);
                if (!selectedChannels.has(idx)) cb.checked = false;
            });

            updateSelectedCount();
        }

        function downloadTemplate() {
            if (selectedChannels.size === 0) {
                showMessage('Please select at least one track', 'error');
                return;
            }

            pendingDownloadType = 'template';
            document.getElementById('filenameModal').querySelector('h2').textContent = 'Save Track Template';
            // Show filename modal
            document.getElementById('filenameModal').classList.add('active');
            document.getElementById('filenameInput').focus();
            document.getElementById('filenameInput').select();
        }

        function downloadCSV() {
            if (selectedChannels.size === 0) {
                showMessage('Please select at least one track', 'error');
                return;
            }

            pendingDownloadType = 'csv';
            document.getElementById('filenameModal').querySelector('h2').textContent = 'Export as CSV';
            document.getElementById('filenameModal').classList.add('active');
            document.getElementById('filenameInput').focus();
            document.getElementById('filenameInput').select();
        }
        
        function closeFilenameModal() {
            document.getElementById('filenameModal').classList.remove('active');
        }
        
        function confirmDownload() {
            const filename = document.getElementById('filenameInput').value.trim();

            if (!filename) {
                alert('Please enter a filename');
                return;
            }

            closeFilenameModal();

            // Get selected channels from current combined list
            const selected = [];
            currentCombinedChannels.forEach((ch, idx) => {
                if (selectedChannels.has(idx)) {
                    selected.push(ch);
                }
            });

            if (pendingDownloadType === 'csv') {
                // Build CSV in browser
                const typeLabels = { inputs: 'Input', aux: 'Aux', groups: 'Group', matrix: 'Matrix' };
                const rows = [['Type', 'Number', 'Name']];
                selected.forEach(ch => {
                    const type = typeLabels[ch.type] || ch.type;
                    const name = ch.name.includes(',') ? '"' + ch.name + '"' : ch.name;
                    rows.push([type, ch.number, name]);
                });
                const csvContent = rows.map(r => r.join(',')).join('\\n');
                const blob = new Blob([csvContent], { type: 'text/csv' });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = filename + '.csv';
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
                showMessage(`✓ CSV "${filename}.csv" with ${selected.length} channels exported!`, 'success');
                return;
            }

            // Request template generation
            fetch('/generate', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ channels: selected, stereo_mode: stereoMode })
            })
            .then(r => r.json())
            .then(data => {
                if (data.success) {
                    // Create blob and download
                    const blob = new Blob([data.template], { type: 'text/plain' });
                    const url = URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = filename + '.RTrackTemplate';
                    document.body.appendChild(a);
                    a.click();
                    document.body.removeChild(a);
                    URL.revokeObjectURL(url);

                    showMessage(`✓ Template "${filename}.RTrackTemplate" with ${selected.length} tracks downloaded!`, 'success');
                } else {
                    showMessage('✗ Error generating template', 'error');
                }
            })
            .catch(err => {
                showMessage('✗ Error: ' + err, 'error');
            });
        }
        
        function toggleDarkMode() {
            const isDark = document.body.classList.toggle('dark');
            document.getElementById('darkModeBtn').textContent = isDark ? '☀️ Light Mode' : '🌙 Dark Mode';
        }

        function setStereoMode(mode) {
            stereoMode = mode;
            ['optSplit', 'optSplitBottom'].forEach(id => {
                const el = document.getElementById(id);
                if (el) el.classList.toggle('active', mode === 'split');
            });
            ['optStereo', 'optStereoBottom'].forEach(id => {
                const el = document.getElementById(id);
                if (el) el.classList.toggle('active', mode === 'stereo');
            });
        }

        // Keyboard shortcuts on track list
        document.addEventListener('keydown', (e) => {
            if (document.activeElement && document.activeElement.tagName === 'INPUT') return;
            if (document.activeElement && document.activeElement.tagName === 'TEXTAREA') return;

            // Cmd/Ctrl+A — select all
            if ((e.metaKey || e.ctrlKey) && e.key === 'a') {
                if (currentCombinedChannels.length === 0) return;
                e.preventDefault();
                activeChannels.clear();
                currentCombinedChannels.forEach((_, i) => activeChannels.add(i));
                lastActiveIdx = currentCombinedChannels.length - 1;
                updateBulkBar();
                document.querySelectorAll('.track-item').forEach(el => el.classList.add('active-highlight'));
            }

            // Delete / Backspace — remove highlighted channels
            if ((e.key === 'Delete' || e.key === 'Backspace') && activeChannels.size > 0) {
                e.preventDefault();
                deleteSelectedChannels();
            }

            // Cmd/Ctrl+Z — undo
            if ((e.metaKey || e.ctrlKey) && e.key === 'z') {
                e.preventDefault();
                undo();
            }
        });

        // Allow Enter key to confirm in modals
        document.getElementById('filenameInput').addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
                confirmDownload();
            }
        });

        document.getElementById('newChannelName').addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
                addCustomChannel();
            }
        });
        
        function applyColorToSection(type, color) {
            const idMap = { inputs: 'colorInputs', aux: 'colorAux', groups: 'colorGroups', matrix: 'colorMatrix' };
            const swatchIdMap = { inputs: 'swatchInputs', aux: 'swatchAux', groups: 'swatchGroups', matrix: 'swatchMatrix' };
            const swatch = document.getElementById(swatchIdMap[type]);
            if (swatch) {
                swatch.style.backgroundColor = color;
                swatch.classList.add('has-color');
            }
            currentCombinedChannels.forEach(ch => {
                if (ch.type === type) ch.color = color;
            });
            showPreview(currentCombinedChannels);
        }

        function clearColorFromSection(type) {
            const swatchIdMap = { inputs: 'swatchInputs', aux: 'swatchAux', groups: 'swatchGroups', matrix: 'swatchMatrix' };
            const colorIdMap = { inputs: 'colorInputs', aux: 'colorAux', groups: 'colorGroups', matrix: 'colorMatrix' };
            const swatch = document.getElementById(swatchIdMap[type]);
            if (swatch) {
                swatch.style.backgroundColor = '';
                swatch.classList.remove('has-color');
                document.getElementById(colorIdMap[type]).value = '#ff6b6b';
            }
            currentCombinedChannels.forEach(ch => {
                if (ch.type === type) ch.color = null;
            });
            showPreview(currentCombinedChannels);
        }

        function handleRowClick(e, idx) {
            if (e.metaKey || e.ctrlKey) {
                // Toggle individual
                if (activeChannels.has(idx)) activeChannels.delete(idx);
                else activeChannels.add(idx);
                lastActiveIdx = idx;
            } else if (e.shiftKey && lastActiveIdx !== null) {
                // Range select
                const start = Math.min(lastActiveIdx, idx);
                const end = Math.max(lastActiveIdx, idx);
                for (let i = start; i <= end; i++) activeChannels.add(i);
            } else {
                // Single select (clear others)
                activeChannels.clear();
                activeChannels.add(idx);
                lastActiveIdx = idx;
            }
            updateBulkBar();
            // Refresh just the highlight classes without full re-render
            document.querySelectorAll('.track-item').forEach((el, i) => {
                el.classList.toggle('active-highlight', activeChannels.has(parseInt(el.dataset.index)));
            });
        }

        function updateBulkBar() {
            const bar = document.getElementById('bulkBar');
            const count = document.getElementById('bulkCount');
            if (activeChannels.size > 1) {
                bar.classList.add('visible');
                count.textContent = activeChannels.size + ' channel' + (activeChannels.size > 1 ? 's' : '') + ' selected';
            } else {
                bar.classList.remove('visible');
            }
        }

        function clearActiveSelection() {
            activeChannels.clear();
            lastActiveIdx = null;
            updateBulkBar();
            document.querySelectorAll('.track-item').forEach(el => el.classList.remove('active-highlight'));
        }

        function clearBulkColors() {
            activeChannels.forEach(i => { currentCombinedChannels[i].color = null; });
            showPreview(currentCombinedChannels);
            updateBulkBar();
        }

        function refreshSectionCounts() {
            const counts = { inputs: 0, aux: 0, groups: 0, matrix: 0 };
            currentCombinedChannels.forEach(ch => {
                if (counts[ch.type] !== undefined) counts[ch.type]++;
            });
            document.getElementById('inputsCount').textContent = counts.inputs;
            document.getElementById('auxCount').textContent = counts.aux;
            document.getElementById('groupsCount').textContent = counts.groups;
            document.getElementById('matrixCount').textContent = counts.matrix;
        }

        function startEditName(idx, nameSpan, ch) {
            const input = document.createElement('input');
            input.type = 'text';
            input.value = ch.name;
            input.style.cssText = 'flex: 1; padding: 2px 6px; border: 1px solid #007aff; border-radius: 4px; font-size: 13px; font-family: inherit; outline: none;';

            let saved = false;
            function save() {
                if (saved) return;
                saved = true;
                const newName = input.value.trim();
                if (newName && newName !== ch.name) {
                    saveUndo();
                    currentCombinedChannels[idx].name = newName;
                }
                showPreview(currentCombinedChannels);
            }

            input.addEventListener('blur', save);
            input.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') { e.preventDefault(); input.blur(); }
                if (e.key === 'Escape') { saved = true; showPreview(currentCombinedChannels); }
                if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
                    e.preventDefault();
                    const newName = input.value.trim();
                    if (newName) currentCombinedChannels[idx].name = newName;
                    saved = true;
                    const nextIdx = e.key === 'ArrowDown' ? idx + 1 : idx - 1;
                    if (nextIdx >= 0 && nextIdx < currentCombinedChannels.length) {
                        showPreview(currentCombinedChannels);
                        // After re-render, find the next name span and start editing it
                        const items = document.querySelectorAll('.track-item');
                        const nextItem = items[nextIdx];
                        if (nextItem) {
                            const nextName = nextItem.querySelector('.track-name');
                            if (nextName) startEditName(nextIdx, nextName, currentCombinedChannels[nextIdx]);
                        }
                    } else {
                        showPreview(currentCombinedChannels);
                    }
                }
            });

            nameSpan.replaceWith(input);
            input.focus();
            input.select();
        }

        function openAddChannelModal() {
            document.getElementById('addChannelModal').classList.add('active');
            document.getElementById('newChannelName').focus();
        }

        function closeAddChannelModal() {
            document.getElementById('addChannelModal').classList.remove('active');
            document.getElementById('newChannelName').value = '';
            document.getElementById('newChannelStereo').checked = false;
            document.getElementById('newChannelType').value = 'inputs';
            document.getElementById('newChannelQty').value = '1';
        }

        function addCustomChannel() {
            const baseName = document.getElementById('newChannelName').value.trim();
            if (!baseName) {
                alert('Please enter a channel name');
                return;
            }

            const type = document.getElementById('newChannelType').value;
            const isStereo = document.getElementById('newChannelStereo').checked;
            const qty = Math.max(1, parseInt(document.getElementById('newChannelQty').value) || 1);

            saveUndo();
            for (let i = 0; i < qty; i++) {
                customChannelCount++;
                const name = qty > 1 ? `${baseName} ${i + 1}` : baseName;
                const number = isStereo ? 'C' + customChannelCount + 's' : 'C' + customChannelCount;
                currentCombinedChannels.push({ number, name, type, isCustom: true });
                selectedChannels.add(currentCombinedChannels.length - 1);
            }

            closeAddChannelModal();
            showPreview(currentCombinedChannels);
            refreshSectionCounts();
            document.getElementById('previewArea').style.display = 'block';
        }

        function deleteChannelWithConfirm(idx) {
            const ch = currentCombinedChannels[idx];
            if (!confirm(`Remove "${ch.name}"?`)) return;
            saveUndo();
            deleteChannel(idx);
        }

        function deleteSelectedChannels() {
            if (activeChannels.size === 0) return;
            const names = Array.from(activeChannels).map(i => currentCombinedChannels[i].name);
            const label = names.length === 1
                ? `Remove "${names[0]}"?`
                : `Remove ${names.length} selected channels?`;
            if (!confirm(label)) return;
            saveUndo();
            // Delete highest indices first to avoid shifting
            const indices = Array.from(activeChannels).sort((a, b) => b - a);
            indices.forEach(i => deleteChannel(i));
        }

        function deleteChannel(idx) {
            const newSelected = new Set();
            selectedChannels.forEach(i => {
                if (i < idx) newSelected.add(i);
                else if (i > idx) newSelected.add(i - 1);
            });
            selectedChannels = newSelected;
            const newActive = new Set();
            activeChannels.forEach(i => {
                if (i < idx) newActive.add(i);
                else if (i > idx) newActive.add(i - 1);
            });
            activeChannels = newActive;
            currentCombinedChannels.splice(idx, 1);
            showPreview(currentCombinedChannels);
            refreshSectionCounts();
            updateBulkBar();
        }

        function reset() {
            const hasChannels = currentCombinedChannels.length > 0;
            if (hasChannels && !confirm('Replace the current session with a new file? This will clear all channels in this tab.')) return;
            document.getElementById('fileInput').value = '';
            document.getElementById('previewArea').style.display = 'none';
            document.getElementById('message').style.display = 'none';
            const fresh = newSessionState(getTab().name);
            sessions[activeTab] = fresh;
            loadTabState();
            updateBulkBar();
            updateUndoBtn();
            // Open file picker immediately
            document.getElementById('fileInput').click();
        }
        
        function showMessage(msg, type) {
            const msgDiv = document.getElementById('message');
            msgDiv.textContent = msg;
            msgDiv.className = 'message ' + type;
            msgDiv.style.display = '';
        }
        
        // Heartbeat check to detect server disconnect
        let heartbeatInterval;
        let missedHeartbeats = 0;
        
        function checkHeartbeat() {
            fetch('/heartbeat', { 
                method: 'GET',
                cache: 'no-cache'
            })
            .then(response => {
                if (response.ok) {
                    missedHeartbeats = 0;
                } else {
                    missedHeartbeats++;
                }
            })
            .catch(err => {
                missedHeartbeats++;
                if (missedHeartbeats >= 2) {
                    // Server is down
                    clearInterval(heartbeatInterval);
                    document.getElementById('disconnectOverlay').style.display = 'flex';
                }
            });
        }
        
        // Check every 3 seconds
        heartbeatInterval = setInterval(checkHeartbeat, 3000);
    </script>
</body>
</html>
        '''
        
        self.send_response(200)
        self.send_header('Content-type', 'text/html')
        self.end_headers()
        self.wfile.write(html.encode())
    
    def handle_conversion(self):
        """Handle file upload and parsing"""
        try:
            # Parse multipart form data
            content_length = int(self.headers['Content-Length'])
            body = self.rfile.read(content_length)

            # Simple multipart parsing (for single file upload)
            boundary = self.headers['Content-Type'].split('boundary=')[1]
            parts = body.split(('--' + boundary).encode())

            file_content = None
            filename = ''
            for part in parts:
                if b'filename=' in part:
                    # Extract filename
                    fn_match = re.search(rb'filename="([^"]+)"', part)
                    if fn_match:
                        filename = fn_match.group(1).decode('utf-8', errors='replace').lower()
                    file_start = part.find(b'\r\n\r\n') + 4
                    file_end = part.rfind(b'\r\n')
                    file_content = part[file_start:file_end]
                    break

            if not file_content:
                self.send_json({'success': False, 'error': 'No file found in upload'})
                return

            # Dispatch to the correct parser based on file extension
            if filename.endswith('.rivagepm'):
                parsed_data = parse_rivage_pm_show_file(file_content)
            elif filename.endswith('.tar.gz'):
                parsed_data = parse_dlive_show_file(file_content)
            else:
                parsed_data = parse_digico_rtf(file_content)
            
            self.send_json({
                'success': True,
                'sections': parsed_data,  # Send all sections
                'counts': {
                    'inputs': len(parsed_data['inputs']),
                    'aux': len(parsed_data['aux']),
                    'groups': len(parsed_data['groups']),
                    'matrix': len(parsed_data['matrix'])
                }
            })
            
        except Exception as e:
            self.send_json({'success': False, 'error': str(e)})
    
    def handle_generate(self):
        """Generate template from selected channels"""
        try:
            content_length = int(self.headers['Content-Length'])
            body = self.rfile.read(content_length)
            data = json.loads(body.decode())
            
            channels = data.get('channels', [])
            stereo_mode = data.get('stereo_mode', 'split')

            if not channels:
                self.send_json({'success': False, 'error': 'No channels selected'})
                return

            # Generate Reaper template
            template = generate_reaper_track_template(channels, stereo_mode)
            
            self.send_json({
                'success': True,
                'template': template,
                'count': len(channels)
            })
            
        except Exception as e:
            self.send_json({'success': False, 'error': str(e)})
    
    def send_json(self, data):
        """Send JSON response"""
        self.send_response(200)
        self.send_header('Content-type', 'application/json')
        self.end_headers()
        self.wfile.write(json.dumps(data).encode())


class DiGiCoConverterApp(rumps.App):
    def __init__(self):
        super(DiGiCoConverterApp, self).__init__("🎛️", quit_button=None)
        self.server = None
        self.server_thread = None
        self.port = None
        
        # Menu items
        self.menu = [
            rumps.MenuItem("Open Converter", callback=self.open_browser),
            rumps.separator,
            rumps.MenuItem("Restart Server", callback=self.restart_server),
            rumps.separator,
            rumps.MenuItem("Quit", callback=self.quit_app)
        ]
        
        # Start server
        self.start_server()
    
    def start_server(self):
        """Start the HTTP server in a background thread"""
        self.port = find_available_port(8081)
        
        if self.port is None:
            rumps.alert(
                title="DiGiCo to Reaper",
                message="Could not find an available port (8081-8090 all in use).\n\nPlease close other applications and try again.",
                ok="Quit"
            )
            rumps.quit_application()
            return
        
        self.server = HTTPServer(('localhost', self.port), DiGiCoToReaperHandler)
        
        # Run server in background thread
        self.server_thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.server_thread.start()
        
        # Update menu title with port
        self.title = f"🎛️ :{self.port}"
        
        # Update menu item
        self.menu["Open Converter"].title = f"Open Converter (:{self.port})"
        
        print(f"✅ Server running on http://localhost:{self.port}")
        
        # Auto-open browser on first launch
        self.open_browser(None)
    
    def open_browser(self, _):
        """Open the converter in default browser"""
        if self.port:
            subprocess.Popen(['open', f'http://localhost:{self.port}'])
    
    def restart_server(self, _):
        """Restart the server"""
        if self.server:
            self.server.shutdown()
            self.server_thread.join(timeout=2)
        
        rumps.notification(
            title="DiGiCo to Reaper",
            subtitle="Restarting server...",
            message=""
        )
        
        self.start_server()
        
        rumps.notification(
            title="DiGiCo to Reaper",
            subtitle="Server restarted",
            message=f"Running on port {self.port}"
        )
    
    def quit_app(self, _):
        """Quit the application"""
        if self.server:
            self.server.shutdown()
        rumps.quit_application()


if __name__ == "__main__":
    DiGiCoConverterApp().run()
