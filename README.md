# PowerPoint Modifier

A flexible tool that modifies text in PowerPoint presentations (both `.ppt` and `.pptx`) based on text replacements defined in a JSON config file. Available as both a **GUI application** and **CLI tool**.

## Features

- ✅ Supports both `.ppt` (legacy) and `.pptx` (modern) formats
- ✅ Text replacement across all slides
- ✅ Handles text in shapes, text boxes, and tables
- ✅ Configurable via JSON config file
- ✅ **User-friendly GUI** with drag-and-drop interface
- ✅ **Command-line interface** for automation
- ✅ **Standalone executables** - no Python installation required!

## Download & Use (No Installation Required)

### For End Users

#### GUI Application (Recommended)

1. **Download** `pptmod-gui.exe` from the releases
2. **Double-click** to launch the application
3. **Select** your PowerPoint file, config file, and output location
4. **Click** "Process PowerPoint" and you're done!

#### Command-Line Tool

1. **Download** `pptmod.exe` from the releases
2. **Create** a `config.json` file with your text replacements (see [Configuration File](#configuration-file))
3. **Run** the tool:
   ```cmd
   pptmod.exe presentation.pptx
   ```

That's it! No Python or dependencies needed.

## Building from Source

### For Developers

#### Install Dependencies

```bash
# Install runtime dependencies
pip install python-pptx pywin32 wxPython

# Install build tools (optional, for creating executable)
pip install pyinstaller

# Or using uv (if you have it)
uv sync
```

#### Build Standalone Executables

```bash
# Build GUI executable
.\build-gui.ps1

# Build CLI executable
.\build.ps1

# Or manually with PyInstaller
pyinstaller pptmod-gui.spec  # For GUI
pyinstaller pptmod.spec      # For CLI
```

The executables will be created in the `dist\` folder.

**Note:** `.ppt` file support requires Microsoft PowerPoint to be installed on Windows.

## Usage

### Using the GUI Application (Easiest)

1. **Launch** `pptmod-gui.exe` or run `python gui.py`
2. **Browse** to select your input PowerPoint file
3. **Optional:** Choose a custom output file location
4. **Optional:** Select a different config file (default: `config.json`)
5. **Click** "Process PowerPoint"
6. **View** the output log for results

The GUI provides:
- File browser dialogs for easy file selection
- Real-time processing log
- Success notifications
- Option to open output folder after completion

### Using the Command-Line Tool

#### Executable (Recommended for End Users)

```cmd
# Basic usage
pptmod.exe presentation.pptx

# Specify custom output file
pptmod.exe input.pptx -o output.pptx

# Use a different config file
pptmod.exe input.pptx -c custom_config.json

# Combine options
pptmod.exe template.ppt -o final.ppt -c replacements.json
```

#### Python Script (For Development)

```bash
# Run CLI
python main.py presentation.pptx

# Run GUI
python gui.py

# Or with uv
uv run main.py presentation.pptx  # CLI
uv run gui.py                      # GUI

# Or with entry points (after uv sync)
pptmod presentation.pptx           # CLI
pptmod-gui                         # GUI
```

This will:
- Read replacements from `config.json` (default)
- Create a modified file named `presentation_modified.pptx`

### Command-Line Options (CLI only)

- `input` - Input PowerPoint file (.ppt or .pptx) - **required**
- `-o, --output` - Output file path (default: `input_modified.ext`)
- `-c, --config` - Config file with text replacements (default: `config.json`)

## Configuration File

Create a JSON file with your text replacements:

```json
{
  "replacements": {
    "{{NAME}}": "John Doe",
    "{{COMPANY}}": "Acme Corporation",
    "{{DATE}}": "November 29, 2025",
    "{{TITLE}}": "Senior Developer",
    "placeholder_text": "actual_text"
  }
}
```

The tool will find all occurrences of the keys (e.g., `{{NAME}}`) and replace them with the corresponding values (e.g., `John Doe`).

## Examples

### Example 1: Simple Template Replacement

```bash
python main.py template.pptx
```

Uses `config.json` to replace template variables.

### Example 2: Multiple Presentations

```bash
python main.py presentation1.pptx -o output1.pptx -c config.json
python main.py presentation2.pptx -o output2.pptx -c config.json
```

### Example 3: Legacy .ppt Files

```bash
python main.py old_template.ppt -o filled_template.ppt
```

**Note:** Requires PowerPoint installed on Windows.

## File Format Support

| Format | Extension | Support | Requirements |
|--------|-----------|---------|--------------|
| Modern PowerPoint | `.pptx` | ✅ Full support | None (bundled in executable) |
| Legacy PowerPoint | `.ppt` | ✅ Full support | PowerPoint installed (Windows only) |

## Distribution

### Sharing with Others

To share the tool with users who don't have Python installed:

#### GUI Application (Easiest for End Users)

1. **Build the GUI executable** (see [Building from Source](#building-from-source))
2. **Package for distribution:**
   - `dist\pptmod-gui.exe` - The standalone GUI application
   - `config.json` - Example configuration file
   - `README.txt` - Simple usage instructions

3. **User Instructions:**
   - Double-click `pptmod-gui.exe` to launch
   - Use the interface to select files and process presentations

#### CLI Tool (For Advanced Users/Automation)

1. **Build the CLI executable** (see [Building from Source](#building-from-source))
2. **Package for distribution:**
   - `dist\pptmod.exe` - The standalone CLI executable
   - `config.json` - Example configuration file
   - `README.txt` - Simple usage instructions

3. **User Instructions:**
   ```cmd
   # Place pptmod.exe and config.json in the same folder
   # Edit config.json with your text replacements
   # Run from command prompt or PowerShell:
   pptmod.exe your_presentation.pptx
   ```

## How It Works

1. **Load Config:** Reads the JSON config file with text replacements
2. **Open Presentation:** Opens the PowerPoint file based on its format
3. **Process Slides:** Iterates through all slides, shapes, and tables
4. **Replace Text:** Finds and replaces all occurrences of the specified text
5. **Save Output:** Saves the modified presentation to the output file

## Troubleshooting

### Error: Config file not found
Make sure `config.json` exists in the current directory or specify a custom path with `-c`.

### Error: .ppt files require PowerPoint
Legacy `.ppt` files use COM automation which requires Microsoft PowerPoint to be installed on Windows.

### No replacements made
Check that:
- Your config file has the correct format
- The text you're searching for exists in the presentation
- Text isn't split across multiple runs (try copying and pasting fresh text)

## License

MIT License - feel free to use and modify as needed!
