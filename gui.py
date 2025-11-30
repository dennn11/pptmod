#!/usr/bin/env python3
"""
PowerPoint Text Modifier GUI
A simple graphical interface for modifying text in PowerPoint presentations.
"""

import wx
import json
import os
from pathlib import Path
from typing import Dict
import threading
from main import modify_pptx, modify_ppt


class PPTModifierFrame(wx.Frame):
    """Main application frame for PPT Modifier."""
    
    def __init__(self):
        super().__init__(
            parent=None,
            title='PowerPoint Text Modifier',
            size=(700, 550)
        )
        
        # Initialize variables
        self.input_file = ""
        self.output_file = ""
        self.config_file = ""
        
        # Create UI
        self.init_ui()
        
        # Center the window
        self.Centre()
        
    def init_ui(self):
        """Initialize the user interface."""
        self.panel = wx.Panel(self)
        main_sizer = wx.BoxSizer(wx.VERTICAL)
        
        # Title
        title = wx.StaticText(self.panel, label="PowerPoint Text Modifier")
        title_font = title.GetFont()
        title_font.PointSize += 4
        title_font = title_font.Bold()
        title.SetFont(title_font)
        main_sizer.Add(title, 0, wx.ALL | wx.CENTER, 10)
        
        # Separator
        main_sizer.Add(wx.StaticLine(self.panel), 0, wx.EXPAND | wx.ALL, 5)
        
        # Input file section
        input_box = wx.StaticBoxSizer(wx.VERTICAL, self.panel, "Input PowerPoint File")
        input_panel = wx.Panel(self.panel)
        input_sizer = wx.BoxSizer(wx.HORIZONTAL)
        
        self.input_text = wx.TextCtrl(input_panel, style=wx.TE_READONLY)
        input_sizer.Add(self.input_text, 1, wx.EXPAND | wx.RIGHT, 5)
        
        browse_input_btn = wx.Button(input_panel, label="Browse...")
        browse_input_btn.Bind(wx.EVT_BUTTON, self.on_browse_input)
        input_sizer.Add(browse_input_btn, 0)
        
        input_panel.SetSizer(input_sizer)
        input_box.Add(input_panel, 0, wx.EXPAND | wx.ALL, 5)
        main_sizer.Add(input_box, 0, wx.EXPAND | wx.ALL, 10)
        
        # Output file section
        output_box = wx.StaticBoxSizer(wx.VERTICAL, self.panel, "Output File (Optional)")
        output_panel = wx.Panel(self.panel)
        output_sizer = wx.BoxSizer(wx.HORIZONTAL)
        
        self.output_text = wx.TextCtrl(output_panel, style=wx.TE_READONLY)
        output_sizer.Add(self.output_text, 1, wx.EXPAND | wx.RIGHT, 5)
        
        browse_output_btn = wx.Button(output_panel, label="Browse...")
        browse_output_btn.Bind(wx.EVT_BUTTON, self.on_browse_output)
        output_sizer.Add(browse_output_btn, 0, wx.RIGHT, 5)
        
        clear_output_btn = wx.Button(output_panel, label="Clear")
        clear_output_btn.Bind(wx.EVT_BUTTON, self.on_clear_output)
        output_sizer.Add(clear_output_btn, 0)
        
        output_panel.SetSizer(output_sizer)
        output_box.Add(output_panel, 0, wx.EXPAND | wx.ALL, 5)
        
        output_note = wx.StaticText(self.panel, label="Leave empty to auto-generate as 'filename_modified.ext'")
        output_note.SetForegroundColour(wx.Colour(100, 100, 100))
        output_box.Add(output_note, 0, wx.LEFT, 5)
        
        main_sizer.Add(output_box, 0, wx.EXPAND | wx.ALL, 10)
        
        # Config file section
        config_box = wx.StaticBoxSizer(wx.VERTICAL, self.panel, "Configuration File")
        config_panel = wx.Panel(self.panel)
        config_sizer = wx.BoxSizer(wx.HORIZONTAL)
        
        self.config_text = wx.TextCtrl(config_panel, value="config.json", style=wx.TE_READONLY)
        config_sizer.Add(self.config_text, 1, wx.EXPAND | wx.RIGHT, 5)
        
        browse_config_btn = wx.Button(config_panel, label="Browse...")
        browse_config_btn.Bind(wx.EVT_BUTTON, self.on_browse_config)
        config_sizer.Add(browse_config_btn, 0)
        
        config_panel.SetSizer(config_sizer)
        config_box.Add(config_panel, 0, wx.EXPAND | wx.ALL, 5)
        main_sizer.Add(config_box, 0, wx.EXPAND | wx.ALL, 10)
        
        # Log/Output section
        log_box = wx.StaticBoxSizer(wx.VERTICAL, self.panel, "Output Log")
        self.log_text = wx.TextCtrl(
            self.panel,
            style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_WORDWRAP,
            size=(-1, 150)
        )
        log_box.Add(self.log_text, 1, wx.EXPAND | wx.ALL, 5)
        self.log_box_sizer = log_box
        self.main_sizer = main_sizer
        main_sizer.Add(log_box, 1, wx.EXPAND | wx.ALL, 10)
        
        # Buttons section
        button_sizer = wx.BoxSizer(wx.HORIZONTAL)
        
        self.process_btn = wx.Button(self.panel, label="Process PowerPoint", size=(150, 35))
        self.process_btn.Bind(wx.EVT_BUTTON, self.on_process)
        button_sizer.Add(self.process_btn, 0, wx.RIGHT, 10)
        
        self.toggle_log_btn = wx.Button(self.panel, label="Hide Log")
        self.toggle_log_btn.Bind(wx.EVT_BUTTON, self.on_toggle_log)
        button_sizer.Add(self.toggle_log_btn, 0, wx.RIGHT, 10)
        
        clear_log_btn = wx.Button(self.panel, label="Clear Log")
        clear_log_btn.Bind(wx.EVT_BUTTON, self.on_clear_log)
        button_sizer.Add(clear_log_btn, 0, wx.RIGHT, 10)
        
        quit_btn = wx.Button(self.panel, label="Quit")
        quit_btn.Bind(wx.EVT_BUTTON, self.on_quit)
        button_sizer.Add(quit_btn, 0)
        
        main_sizer.Add(button_sizer, 0, wx.ALL | wx.CENTER, 10)
        
        self.panel.SetSizer(main_sizer)
        
        # Track log visibility
        self.log_visible = False
        self.hide_log()
        
        # Set default config file
        self.config_file = "config.json"
        
    def on_toggle_log(self, event):
        """Toggle the visibility of the log section."""
        if self.log_visible:
            self.hide_log()
        else:
            self.show_log()
    
    def show_log(self):
        """Show the log section."""
        self.main_sizer.Show(self.log_box_sizer, True)
        self.toggle_log_btn.SetLabel("Hide Log")
        self.log_visible = True
        self.panel.Layout()
        # Resize window to accommodate log
        current_size = self.GetSize()
        self.SetSize((current_size.width, 550))

    def hide_log(self):
        """Hide the log section."""
        self.main_sizer.Show(self.log_box_sizer, False)
        self.toggle_log_btn.SetLabel("Show Log")
        self.log_visible = False
        self.panel.Layout()
        # Resize window to be smaller without log
        current_size = self.GetSize()
        self.SetSize((current_size.width, 400))
        self.Fit()

    def on_browse_input(self, event):
        """Handle browse input file button."""
        wildcard = "PowerPoint files (*.ppt;*.pptx)|*.ppt;*.pptx|All files (*.*)|*.*"
        dialog = wx.FileDialog(
            self,
            "Select Input PowerPoint File",
            wildcard=wildcard,
            style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST
        )
        
        if dialog.ShowModal() == wx.ID_OK:
            self.input_file = dialog.GetPath()
            self.input_text.SetValue(self.input_file)
            self.log(f"Selected input file: {self.input_file}")
            
            # Auto-generate output filename if not set
            if not self.output_file:
                input_path = Path(self.input_file)
                auto_output = str(input_path.parent / f"{input_path.stem}_modified{input_path.suffix}")
                self.output_text.SetValue(auto_output)
        
        dialog.Destroy()
        
    def on_browse_output(self, event):
        """Handle browse output file button."""
        wildcard = "PowerPoint files (*.ppt;*.pptx)|*.ppt;*.pptx|All files (*.*)|*.*"
        dialog = wx.FileDialog(
            self,
            "Select Output PowerPoint File",
            wildcard=wildcard,
            style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT
        )
        
        if dialog.ShowModal() == wx.ID_OK:
            self.output_file = dialog.GetPath()
            self.output_text.SetValue(self.output_file)
            self.log(f"Selected output file: {self.output_file}")
        
        dialog.Destroy()
        
    def on_clear_output(self, event):
        """Clear the output file selection."""
        self.output_file = ""
        self.output_text.SetValue("")
        self.log("Output file cleared. Will auto-generate filename.")
        
    def on_browse_config(self, event):
        """Handle browse config file button."""
        wildcard = "JSON files (*.json)|*.json|All files (*.*)|*.*"
        dialog = wx.FileDialog(
            self,
            "Select Configuration File",
            wildcard=wildcard,
            style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST
        )
        
        if dialog.ShowModal() == wx.ID_OK:
            self.config_file = dialog.GetPath()
            self.config_text.SetValue(self.config_file)
            self.log(f"Selected config file: {self.config_file}")
        
        dialog.Destroy()
        
    def on_process(self, event):
        """Handle process button click."""
        # Validate inputs
        if not self.input_file:
            wx.MessageBox(
                "Please select an input PowerPoint file.",
                "Missing Input",
                wx.OK | wx.ICON_WARNING
            )
            return
        
        if not os.path.exists(self.input_file):
            wx.MessageBox(
                f"Input file does not exist:\n{self.input_file}",
                "File Not Found",
                wx.OK | wx.ICON_ERROR
            )
            return
        
        if not os.path.exists(self.config_file):
            wx.MessageBox(
                f"Config file does not exist:\n{self.config_file}",
                "Config Not Found",
                wx.OK | wx.ICON_ERROR
            )
            return
        
        # Determine output file
        output_file = self.output_file if self.output_file else self.output_text.GetValue()
        if not output_file:
            input_path = Path(self.input_file)
            output_file = str(input_path.parent / f"{input_path.stem}_modified{input_path.suffix}")
        
        # Disable button during processing
        self.process_btn.Enable(False)
        self.log("\n" + "="*50)
        self.log("Starting processing...")
        self.log(f"Input: {self.input_file}")
        self.log(f"Output: {output_file}")
        self.log(f"Config: {self.config_file}")
        
        # Run processing in a separate thread to keep UI responsive
        thread = threading.Thread(
            target=self.process_file,
            args=(self.input_file, output_file, self.config_file)
        )
        thread.daemon = True
        thread.start()
        
    def process_file(self, input_file: str, output_file: str, config_file: str):
        """Process the PowerPoint file (runs in separate thread)."""
        try:
            # Load config
            wx.CallAfter(self.log, "Loading configuration...")
            with open(config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            if 'replacements' not in config:
                wx.CallAfter(self.log, "ERROR: Config file must contain 'replacements' key")
                wx.CallAfter(self.process_btn.Enable, True)
                return
            
            replacements = config['replacements']
            wx.CallAfter(self.log, f"Loaded {len(replacements)} replacement(s)")
            
            # Determine file type
            file_ext = Path(input_file).suffix.lower()
            
            # Process file
            wx.CallAfter(self.log, "Processing presentation...")
            
            if file_ext == '.pptx':
                modify_pptx(input_file, output_file, replacements)
            elif file_ext == '.ppt':
                modify_ppt(input_file, output_file, replacements)
            else:
                wx.CallAfter(self.log, f"ERROR: Unsupported file type '{file_ext}'")
                wx.CallAfter(self.process_btn.Enable, True)
                return
            
            wx.CallAfter(self.log, "✓ Processing completed successfully!")
            wx.CallAfter(self.log, f"Output saved to: {output_file}")
            wx.CallAfter(self.show_success, output_file)
            
        except Exception as e:
            wx.CallAfter(self.log, f"ERROR: {str(e)}")
            wx.CallAfter(wx.MessageBox, f"Error processing file:\n{str(e)}", "Error", wx.OK | wx.ICON_ERROR)
        
        finally:
            wx.CallAfter(self.process_btn.Enable, True)
    
    def show_success(self, output_file):
        """Show success dialog."""
        result = wx.MessageBox(
            f"PowerPoint file processed successfully!\n\nOutput saved to:\n{output_file}\n\nWould you like to open the output folder?",
            "Success",
            wx.YES_NO | wx.ICON_INFORMATION
        )
        
        if result == wx.YES:
            # Open the folder containing the output file
            folder = str(Path(output_file).parent)
            os.startfile(folder)
    
    def log(self, message: str):
        """Add a message to the log."""
        self.log_text.AppendText(message + "\n")
        
    def on_clear_log(self, event):
        """Clear the log text."""
        self.log_text.Clear()
        
    def on_quit(self, event):
        """Handle quit button."""
        self.Close()


def main():
    """Main entry point for the GUI application."""
    app = wx.App()
    frame = PPTModifierFrame()
    frame.Show()
    app.MainLoop()


if __name__ == "__main__":
    main()
