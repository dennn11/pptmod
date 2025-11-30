#!/usr/bin/env python3
"""
PowerPoint Text Modifier GUI
A simple graphical interface for modifying text in PowerPoint presentations.
"""

import wx
import json
import os
import wx.grid
from pathlib import Path
from typing import Dict
import threading
from main import modify_pptx, modify_ppt, export_to_pdf


class PPTModifierFrame(wx.Frame):
    """Main application frame for PPT Modifier."""
    
    def __init__(self):
        super().__init__(
            parent=None,
            title='pptmod',
            size=(750, 700)
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
        title = wx.StaticText(self.panel, label="pptmod")
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
        
        # Config file section with editable replacements
        config_box = wx.StaticBoxSizer(wx.VERTICAL, self.panel, "Text Replacements")
        
        # Config file selector
        config_panel = wx.Panel(self.panel)
        config_sizer = wx.BoxSizer(wx.HORIZONTAL)
        
        self.config_text = wx.TextCtrl(config_panel, value="config.json", style=wx.TE_READONLY)
        config_sizer.Add(self.config_text, 1, wx.EXPAND | wx.RIGHT, 5)
        
        browse_config_btn = wx.Button(config_panel, label="Load Config")
        browse_config_btn.Bind(wx.EVT_BUTTON, self.on_browse_config)
        config_sizer.Add(browse_config_btn, 0, wx.RIGHT, 5)
        
        save_config_btn = wx.Button(config_panel, label="Save Config")
        save_config_btn.Bind(wx.EVT_BUTTON, self.on_save_config)
        config_sizer.Add(save_config_btn, 0)
        
        config_panel.SetSizer(config_sizer)
        config_box.Add(config_panel, 0, wx.EXPAND | wx.ALL, 5)
        
        # Grid for replacements
        self.replacement_grid = wx.grid.Grid(self.panel)
        self.replacement_grid.CreateGrid(5, 2)
        self.replacement_grid.SetColLabelValue(0, "Find Text")
        self.replacement_grid.SetColLabelValue(1, "Replace With")
        self.replacement_grid.SetColSize(0, 250)
        self.replacement_grid.SetColSize(1, 250)
        self.replacement_grid.SetRowLabelSize(40)
        self.replacement_grid.SetMinSize((-1, 200))
        self.replacement_grid.EnableEditing(True)
        config_box.Add(self.replacement_grid, 1, wx.EXPAND | wx.ALL, 5)
        
        # Grid control buttons
        grid_btn_panel = wx.Panel(self.panel)
        grid_btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        
        add_row_btn = wx.Button(grid_btn_panel, label="Add Row")
        add_row_btn.Bind(wx.EVT_BUTTON, self.on_add_row)
        grid_btn_sizer.Add(add_row_btn, 0, wx.RIGHT, 5)
        
        remove_row_btn = wx.Button(grid_btn_panel, label="Remove Row")
        remove_row_btn.Bind(wx.EVT_BUTTON, self.on_remove_row)
        grid_btn_sizer.Add(remove_row_btn, 0, wx.RIGHT, 5)
        
        clear_all_btn = wx.Button(grid_btn_panel, label="Clear All")
        clear_all_btn.Bind(wx.EVT_BUTTON, self.on_clear_all_rows)
        grid_btn_sizer.Add(clear_all_btn, 0)
        
        grid_btn_panel.SetSizer(grid_btn_sizer)
        config_box.Add(grid_btn_panel, 0, wx.ALL | wx.CENTER, 5)
        
        main_sizer.Add(config_box, 1, wx.EXPAND | wx.ALL, 10)
        
        # Log/Output section
        log_box = wx.StaticBoxSizer(wx.VERTICAL, self.panel, "Output Log")
        self.log_text = wx.TextCtrl(
            self.panel,
            style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_WORDWRAP,
            size=(-1, 100)
        )
        log_box.Add(self.log_text, 1, wx.EXPAND | wx.ALL, 5)
        self.log_box_sizer = log_box
        self.main_sizer = main_sizer
        main_sizer.Add(log_box, 0, wx.EXPAND | wx.ALL, 10)
        
        # Buttons section
        button_sizer = wx.BoxSizer(wx.HORIZONTAL)
        
        self.process_btn = wx.Button(self.panel, label="Process PowerPoint", size=(150, 35))
        self.process_btn.Bind(wx.EVT_BUTTON, self.on_process)
        button_sizer.Add(self.process_btn, 0, wx.RIGHT, 10)
        
        self.export_pdf_btn = wx.Button(self.panel, label="Export to PDF", size=(120, 35))
        self.export_pdf_btn.Bind(wx.EVT_BUTTON, self.on_export_pdf)
        button_sizer.Add(self.export_pdf_btn, 0, wx.RIGHT, 10)
        
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
        
        # Set default config file and load it
        self.config_file = "config.json"
        self.load_config_into_grid(self.config_file)
    
    def load_config_into_grid(self, config_path: str):
        """Load config file into the replacement grid."""
        try:
            if not os.path.exists(config_path):
                self.log(f"Config file not found: {config_path}")
                return
            
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            if 'replacements' not in config:
                self.log("Warning: Config file missing 'replacements' key")
                return
            
            replacements = config['replacements']
            
            # Clear existing grid
            if self.replacement_grid.GetNumberRows() > 0:
                self.replacement_grid.DeleteRows(0, self.replacement_grid.GetNumberRows())
            
            # Add rows for replacements
            num_rows = len(replacements)
            if num_rows > 0:
                self.replacement_grid.AppendRows(num_rows)
                
                row = 0
                for key, value in replacements.items():
                    self.replacement_grid.SetCellValue(row, 0, key)
                    self.replacement_grid.SetCellValue(row, 1, value)
                    row += 1
            
            self.log(f"Loaded {num_rows} replacement(s) from {config_path}")
            
        except Exception as e:
            self.log(f"Error loading config: {str(e)}")
    
    def get_replacements_from_grid(self) -> Dict[str, str]:
        """Get all replacements from the grid."""
        replacements = {}
        num_rows = self.replacement_grid.GetNumberRows()
        
        for row in range(num_rows):
            key = self.replacement_grid.GetCellValue(row, 0).strip()
            value = self.replacement_grid.GetCellValue(row, 1).strip()
            
            # Only add non-empty pairs
            if key:
                replacements[key] = value
        
        return replacements
    
    def on_add_row(self, event):
        """Add a new row to the replacement grid."""
        self.replacement_grid.AppendRows(1)
        self.log("Added new row")
    
    def on_remove_row(self, event):
        """Remove the selected row from the grid."""
        selected_rows = self.replacement_grid.GetSelectedRows()
        
        if not selected_rows:
            # Try to get the current cursor position
            current_row = self.replacement_grid.GetGridCursorRow()
            if current_row >= 0 and self.replacement_grid.GetNumberRows() > 0:
                self.replacement_grid.DeleteRows(current_row, 1)
                self.log(f"Removed row {current_row + 1}")
            else:
                wx.MessageBox("Please select a row to remove", "No Selection", wx.OK | wx.ICON_INFORMATION)
        else:
            # Remove all selected rows (from bottom to top to maintain indices)
            for row in sorted(selected_rows, reverse=True):
                self.replacement_grid.DeleteRows(row, 1)
            self.log(f"Removed {len(selected_rows)} row(s)")
    
    def on_clear_all_rows(self, event):
        """Clear all rows from the grid."""
        result = wx.MessageBox(
            "Are you sure you want to clear all replacements?",
            "Confirm Clear",
            wx.YES_NO | wx.ICON_QUESTION
        )
        
        if result == wx.YES:
            num_rows = self.replacement_grid.GetNumberRows()
            if num_rows > 0:
                self.replacement_grid.DeleteRows(0, num_rows)
            self.log("Cleared all replacements")
    
    def on_save_config(self, event):
        """Save the current replacements to a config file."""
        # Get replacements from grid
        replacements = self.get_replacements_from_grid()
        
        if not replacements:
            wx.MessageBox(
                "No replacements to save. Please add at least one replacement.",
                "Nothing to Save",
                wx.OK | wx.ICON_WARNING
            )
            return
        
        # Ask user where to save
        wildcard = "JSON files (*.json)|*.json|All files (*.*)|*.*"
        dialog = wx.FileDialog(
            self,
            "Save Configuration File",
            defaultFile=self.config_file,
            wildcard=wildcard,
            style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT
        )
        
        if dialog.ShowModal() == wx.ID_OK:
            save_path = dialog.GetPath()
            
            try:
                config_data = {"replacements": replacements}
                
                with open(save_path, 'w', encoding='utf-8') as f:
                    json.dump(config_data, f, indent=2, ensure_ascii=False)
                
                self.config_file = save_path
                self.config_text.SetValue(save_path)
                self.log(f"Saved {len(replacements)} replacement(s) to {save_path}")
                
                wx.MessageBox(
                    f"Configuration saved successfully!\n\n{save_path}",
                    "Success",
                    wx.OK | wx.ICON_INFORMATION
                )
                
            except Exception as e:
                self.log(f"Error saving config: {str(e)}")
                wx.MessageBox(
                    f"Error saving configuration:\n{str(e)}",
                    "Save Error",
                    wx.OK | wx.ICON_ERROR
                )
        
        dialog.Destroy()
        
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
        self.SetSize((current_size.width, 800))

    def hide_log(self):
        """Hide the log section."""
        self.main_sizer.Show(self.log_box_sizer, False)
        self.toggle_log_btn.SetLabel("Show Log")
        self.log_visible = False
        self.panel.Layout()
        # Resize window to be smaller without log
        current_size = self.GetSize()
        self.SetSize((current_size.width, 600))
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
            "Load Configuration File",
            wildcard=wildcard,
            style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST
        )
        
        if dialog.ShowModal() == wx.ID_OK:
            self.config_file = dialog.GetPath()
            self.config_text.SetValue(self.config_file)
            self.load_config_into_grid(self.config_file)
        
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
        
        # Get replacements from grid instead of config file
        replacements = self.get_replacements_from_grid()
        
        if not replacements:
            wx.MessageBox(
                "No replacements defined. Please add at least one replacement in the grid.",
                "No Replacements",
                wx.OK | wx.ICON_WARNING
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
        self.log(f"Replacements: {len(replacements)}")
        
        # Run processing in a separate thread to keep UI responsive
        thread = threading.Thread(
            target=self.process_file,
            args=(self.input_file, output_file, replacements)
        )
        thread.daemon = True
        thread.start()
        
    def process_file(self, input_file: str, output_file: str, replacements: Dict[str, str]):
        """Process the PowerPoint file (runs in separate thread)."""
        try:
            wx.CallAfter(self.log, f"Using {len(replacements)} replacement(s)")
            
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
    
    def on_export_pdf(self, event):
        """Handle export to PDF button click."""
        # Validate input file
        if not self.input_file:
            wx.MessageBox(
                "Please select an input PowerPoint file first.",
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
        
        # Ask user where to save PDF
        input_path = Path(self.input_file)
        default_pdf = str(input_path.parent / f"{input_path.stem}.pdf")
        
        wildcard = "PDF files (*.pdf)|*.pdf|All files (*.*)|*.*"
        dialog = wx.FileDialog(
            self,
            "Save PDF As",
            defaultDir=str(input_path.parent),
            defaultFile=f"{input_path.stem}.pdf",
            wildcard=wildcard,
            style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT
        )
        
        if dialog.ShowModal() != wx.ID_OK:
            dialog.Destroy()
            return
        
        pdf_path = dialog.GetPath()
        dialog.Destroy()
        
        # Disable buttons during export
        self.export_pdf_btn.Enable(False)
        self.process_btn.Enable(False)
        
        self.log("\n" + "="*50)
        self.log("Exporting to PDF...")
        self.log(f"Input: {self.input_file}")
        self.log(f"Output PDF: {pdf_path}")
        
        # Run export in a separate thread
        thread = threading.Thread(
            target=self.export_to_pdf_thread,
            args=(self.input_file, pdf_path)
        )
        thread.daemon = True
        thread.start()
    
    def export_to_pdf_thread(self, input_file: str, pdf_path: str):
        """Export PowerPoint to PDF (runs in separate thread)."""
        try:
            export_to_pdf(input_file, pdf_path)
            wx.CallAfter(self.log, "✓ PDF export completed successfully!")
            wx.CallAfter(self.log, f"PDF saved to: {pdf_path}")
            wx.CallAfter(self.show_pdf_success, pdf_path)
        
        except Exception as e:
            wx.CallAfter(self.log, f"ERROR: {str(e)}")
            wx.CallAfter(wx.MessageBox, f"Error exporting to PDF:\n{str(e)}", "Error", wx.OK | wx.ICON_ERROR)
        
        finally:
            wx.CallAfter(self.export_pdf_btn.Enable, True)
            wx.CallAfter(self.process_btn.Enable, True)
    
    def show_pdf_success(self, pdf_path):
        """Show PDF export success dialog."""
        result = wx.MessageBox(
            f"PDF exported successfully!\n\nSaved to:\n{pdf_path}\n\nWould you like to open the PDF folder?",
            "Success",
            wx.YES_NO | wx.ICON_INFORMATION
        )
        
        if result == wx.YES:
            folder = str(Path(pdf_path).parent)
            os.startfile(folder)
        
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
