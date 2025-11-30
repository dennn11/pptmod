# PowerPoint Modifier CLI

A flexible CLI tool that modifies text in PowerPoint presentations (both `.ppt` and `.pptx`) based on text replacements defined in a JSON config file.

## Features

- ✅ Supports both `.ppt` (legacy) and `.pptx` (modern) formats
- ✅ Text replacement across all slides
- ✅ Handles text in shapes, text boxes, and tables
- ✅ Configurable via JSON config file
- ✅ Simple command-line interface
- ✅ **Standalone executable** - no Python installation required!

## Download & Use (No Installation Required)

### For End Users

1. **Download** the `pptmod.exe` from the releases
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
pip install python-pptx pywin32

# Install build tools (optional, for creating executable)
pip install pyinstaller

# Or using uv (if you have it)
uv pip install -e ".[build]"
```

#### Build Standalone Executable

```bash
# Run the build script
.\build.ps1

# Or manually with PyInstaller
pyinstaller pptmod.spec
```

The executable will be created in `dist\pptmod.exe`.

**Note:** `.ppt` file support requires Microsoft PowerPoint to be installed on Windows.

## Usage

### Using the Executable (Recommended for End Users)

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

### Using Python Script (For Development)

```bash
# Basic usage
python main.py presentation.pptx

# Or with uv
uv run main.py presentation.pptx
```

This will:
- Read replacements from `config.json` (default)
- Create a modified file named `presentation_modified.pptx`

### Command-Line Options

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

1. **Build the executable** (see [Building from Source](#building-from-source))
2. **Package for distribution:**
   - `dist\pptmod.exe` - The standalone executable
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
