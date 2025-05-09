"""
GUI module for the UCO to UDO Reconciliation tool.

This module provides the graphical user interface components
using Tkinter for the reconciliation tool.
"""

import os
import sys
import logging
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
from datetime import datetime
from typing import Any, Optional, List, Tuple, Dict, Callable, Union
import threading
from pathlib import Path
import webbrowser

from src.uco_to_udo_recon.core.excel_operations import (
    copy_and_rename_sheet, create_copy_of_target_file
)
from src.uco_to_udo_recon.core.reconciliation import find_table_range
from src.uco_to_udo_recon.utils.file_utils import ensure_file_handle_release, open_excel_file


class TextHandler(logging.Handler):
    """
    Custom logging handler that writes to a Tkinter text widget.
    """
    
    def __init__(self, text_widget: tk.Text):
        """
        Initialize the handler with a text widget.
        
        Args:
            text_widget: The Tkinter text widget to write logs to
        """
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record: logging.LogRecord) -> None:
        """
        Emit a log record to the text widget.
        
        Args:
            record: The log record to emit
            
        Returns:
            None
        """
        msg = self.format(record)
        
        # Use tags to color-code different log levels
        tag = None
        if record.levelno >= logging.ERROR:
            tag = "error"
        elif record.levelno >= logging.WARNING:
            tag = "warning"
        elif record.levelno >= logging.INFO:
            tag = "info"
        elif record.levelno >= logging.DEBUG:
            tag = "debug"
            
        self.text_widget.insert(tk.END, msg + '\n', tag)
        self.text_widget.see(tk.END)
        self.text_widget.update_idletasks()


class HyperlinkLabel(ttk.Label):
    """A ttk.Label that behaves like a hyperlink."""
    
    def __init__(self, master=None, text="", url="", **kwargs):
        super().__init__(master, text=text, cursor="hand2", foreground="#0000FF", **kwargs)
        self.url = url
        self.bind("<Button-1>", self._open_url)
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        
    def _open_url(self, event=None):
        """Open the URL in the default browser."""
        webbrowser.open(self.url)
        
    def _on_enter(self, event=None):
        """Change appearance when mouse enters the label."""
        self.configure(foreground="#000080", font=("TkDefaultFont", 9, "underline"))
        
    def _on_leave(self, event=None):
        """Change appearance when mouse leaves the label."""
        self.configure(foreground="#0000FF", font=("TkDefaultFont", 9, ""))


class FileInputFrame(ttk.LabelFrame):
    """Reusable frame for file input with validation."""
    
    def __init__(self, master, label_text, description=None, filetypes=None, **kwargs):
        """
        Initialize a file input frame.
        
        Args:
            master: Parent widget
            label_text: Label for the file input
            description: Optional description text
            filetypes: List of file types for file dialog
            **kwargs: Additional keyword arguments for the frame
        """
        super().__init__(master, text=label_text, **kwargs)
        self.filetypes = filetypes or [("Excel files", "*.xlsx")]
        self.description = description
        
        self.columnconfigure(0, weight=1)
        
        # Main file input row
        self.file_var = tk.StringVar()
        self.file_var.trace_add("write", self._validate_file_path)
        
        self.file_entry = ttk.Entry(self, textvariable=self.file_var)
        self.file_entry.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        
        self.browse_button = ttk.Button(
            self, 
            text="Browse...", 
            command=self._browse_file
        )
        self.browse_button.grid(row=0, column=1, padx=5, pady=5)
        
        # Add description if provided
        if description:
            desc_label = ttk.Label(
                self, 
                text=description, 
                font=("TkDefaultFont", 8),
                foreground="#777777",
                wraplength=400
            )
            desc_label.grid(row=1, column=0, columnspan=2, sticky="w", padx=5)
            
        # Add validation indicator
        self.status_frame = ttk.Frame(self)
        self.status_frame.grid(row=0, column=2, padx=2, pady=2)
        
        self.status_indicator = ttk.Label(self.status_frame, text="")
        self.status_indicator.pack()
        
        # Initialize validation state
        self.is_valid = False
        self._update_validation_indicator()
        
    def _browse_file(self):
        """Open file dialog and set selected file."""
        filename = filedialog.askopenfilename(filetypes=self.filetypes)
        if filename:
            self.file_var.set(filename)
            
    def _validate_file_path(self, *args):
        """Validate the file path and update status indicator."""
        file_path = self.file_var.get().strip()
        self.is_valid = bool(file_path and os.path.exists(file_path))
        self._update_validation_indicator()
        
    def _update_validation_indicator(self):
        """Update the validation status indicator."""
        if not self.file_var.get().strip():
            self.status_indicator.config(text="⬤", foreground="#999999")  # Gray dot for empty
        elif self.is_valid:
            self.status_indicator.config(text="✓", foreground="#00AA00")  # Green check for valid
        else:
            self.status_indicator.config(text="✗", foreground="#CC0000")  # Red X for invalid
            
    def get_file_path(self):
        """Get the selected file path."""
        return self.file_var.get().strip()
    
    def set_file_path(self, path):
        """Set the file path."""
        self.file_var.set(path)


class SettingsDialog(tk.Toplevel):
    """Dialog for configuring application settings."""
    
    def __init__(self, parent, settings):
        """Initialize the settings dialog."""
        super().__init__(parent)
        self.parent = parent
        self.settings = settings.copy()
        self.result = None

        self.title("Settings")
        self.geometry("500x400")
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()

        self._create_widgets()

        # Center the dialog on the parent window
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = parent.winfo_rootx() + (parent.winfo_width() - width) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - height) // 2
        self.geometry(f"+{x}+{y}")

        # Make dialog modal
        self.protocol("WM_DELETE_WINDOW", self.cancel)
        self.wait_window(self)

    def _create_widgets(self):
        """Create the dialog widgets."""
        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Create notebook for settings categories
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # General settings tab
        general_tab = ttk.Frame(notebook, padding=10)
        notebook.add(general_tab, text="General")

        # Auto-open results setting
        self.auto_open_var = tk.BooleanVar(value=self.settings.get('auto_open_results', True))
        ttk.Checkbutton(
            general_tab,
            text="Automatically open results file when processing completes",
            variable=self.auto_open_var
        ).pack(anchor=tk.W, pady=5)

        # Default component
        ttk.Label(general_tab, text="Default Component:").pack(anchor=tk.W, pady=(10, 2))
        self.default_component_var = tk.StringVar(value=self.settings.get('default_component', "WMD"))
        component_combo = ttk.Combobox(
            general_tab,
            textvariable=self.default_component_var,
            values=["CBP", "CG", "CIS", "CYB", "FEM", "FLE", "ICE", "MGA", "MGT", "OIG", "TSA", "SS", "ST", "WMD"],
            state="readonly",
            width=30
        )
        component_combo.pack(anchor=tk.W, pady=(0, 10))

        # Recent files limit
        ttk.Label(general_tab, text="Number of recent files to remember:").pack(anchor=tk.W, pady=(10, 2))
        self.recent_files_limit_var = tk.IntVar(value=self.settings.get('recent_files_limit', 5))
        recent_files_spin = ttk.Spinbox(
            general_tab,
            from_=0,
            to=20,
            textvariable=self.recent_files_limit_var,
            width=5
        )
        recent_files_spin.pack(anchor=tk.W, pady=(0, 10))

        # Default files location
        ttk.Label(general_tab, text="Default files location:").pack(anchor=tk.W, pady=(10, 2))
        self.default_location_var = tk.StringVar(value=self.settings.get('default_location', ""))
        location_frame = ttk.Frame(general_tab)
        location_frame.pack(fill=tk.X, pady=(0, 10))

        location_entry = ttk.Entry(location_frame, textvariable=self.default_location_var)
        location_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        browse_button = ttk.Button(
            location_frame,
            text="Browse...",
            command=self._browse_default_location
        )
        browse_button.pack(side=tk.RIGHT)

        # Appearance settings tab
        appearance_tab = ttk.Frame(notebook, padding=10)
        notebook.add(appearance_tab, text="Appearance")

        # Theme selection
        ttk.Label(appearance_tab, text="Application Theme:").pack(anchor=tk.W, pady=(5, 2))
        self.theme_var = tk.StringVar(value=self.settings.get('theme', "dark"))

        # Frame for theme preview and selection
        theme_frame = ttk.Frame(appearance_tab)
        theme_frame.pack(fill=tk.X, pady=(5, 15))
        theme_frame.columnconfigure(0, weight=1)
        theme_frame.columnconfigure(1, weight=1)

        # Dark theme option
        dark_frame = ttk.Frame(theme_frame, padding=5)
        dark_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")

        dark_preview = ttk.Frame(dark_frame, style="Preview.Dark.TFrame", height=80, width=120)
        dark_preview.pack(pady=5)

        dark_radio = ttk.Radiobutton(
            dark_frame,
            text="Dark Theme",
            variable=self.theme_var,
            value="dark"
        )
        dark_radio.pack(pady=5)

        # Light theme option
        light_frame = ttk.Frame(theme_frame, padding=5)
        light_frame.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")

        light_preview = ttk.Frame(light_frame, style="Preview.Light.TFrame", height=80, width=120)
        light_preview.pack(pady=5)

        light_radio = ttk.Radiobutton(
            light_frame,
            text="Light Theme",
            variable=self.theme_var,
            value="light"
        )
        light_radio.pack(pady=5)

        # Windows/Default theme option
        system_frame = ttk.Frame(theme_frame, padding=5)
        system_frame.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")

        system_preview = ttk.Frame(system_frame, style="Preview.System.TFrame", height=80, width=120)
        system_preview.pack(pady=5)

        system_radio = ttk.Radiobutton(
            system_frame,
            text="System Default",
            variable=self.theme_var,
            value="system"
        )
        system_radio.pack(pady=5)

        # Blue theme option
        blue_frame = ttk.Frame(theme_frame, padding=5)
        blue_frame.grid(row=1, column=1, padx=5, pady=5, sticky="nsew")

        blue_preview = ttk.Frame(blue_frame, style="Preview.Blue.TFrame", height=80, width=120)
        blue_preview.pack(pady=5)

        blue_radio = ttk.Radiobutton(
            blue_frame,
            text="Blue Theme",
            variable=self.theme_var,
            value="blue"
        )
        blue_radio.pack(pady=5)

        # UI Density setting
        ttk.Separator(appearance_tab, orient="horizontal").pack(fill=tk.X, pady=10)

        ttk.Label(appearance_tab, text="UI Density:").pack(anchor=tk.W, pady=(10, 5))
        self.ui_density_var = tk.StringVar(value=self.settings.get('ui_density', "normal"))

        density_frame = ttk.Frame(appearance_tab)
        density_frame.pack(fill=tk.X)

        ttk.Radiobutton(
            density_frame,
            text="Compact",
            variable=self.ui_density_var,
            value="compact"
        ).pack(side=tk.LEFT, padx=(0, 15))

        ttk.Radiobutton(
            density_frame,
            text="Normal",
            variable=self.ui_density_var,
            value="normal"
        ).pack(side=tk.LEFT, padx=(0, 15))

        ttk.Radiobutton(
            density_frame,
            text="Comfortable",
            variable=self.ui_density_var,
            value="comfortable"
        ).pack(side=tk.LEFT)

        # Advanced settings tab
        advanced_tab = ttk.Frame(notebook, padding=10)
        notebook.add(advanced_tab, text="Advanced")

        # Log level setting
        ttk.Label(advanced_tab, text="Log Level:").pack(anchor=tk.W, pady=(5, 2))
        self.log_level_var = tk.StringVar(value=self.settings.get('log_level', "INFO"))
        log_level_combo = ttk.Combobox(
            advanced_tab,
            textvariable=self.log_level_var,
            values=["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"],
            state="readonly",
            width=10
        )
        log_level_combo.pack(anchor=tk.W, pady=(0, 10))

        # COM Timeout setting
        ttk.Label(advanced_tab, text="Excel COM Operation Timeout (seconds):").pack(anchor=tk.W, pady=(10, 2))
        self.com_timeout_var = tk.IntVar(value=self.settings.get('com_timeout', 30))
        com_timeout_spin = ttk.Spinbox(
            advanced_tab,
            from_=5,
            to=120,
            textvariable=self.com_timeout_var,
            width=5
        )
        com_timeout_spin.pack(anchor=tk.W, pady=(0, 10))

        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, padx=5, pady=10)

        ttk.Button(
            button_frame,
            text="Cancel",
            command=self.cancel
        ).pack(side=tk.RIGHT, padx=5)

        ttk.Button(
            button_frame,
            text="OK",
            command=self.apply_settings
        ).pack(side=tk.RIGHT, padx=5)

        ttk.Button(
            button_frame,
            text="Restore Defaults",
            command=self._restore_defaults
        ).pack(side=tk.LEFT, padx=5)

    def _browse_default_location(self):
        """Browse for default files location."""
        directory = filedialog.askdirectory()
        if directory:
            self.default_location_var.set(directory)

    def _restore_defaults(self):
        """Restore default settings."""
        self.auto_open_var.set(True)
        self.default_component_var.set("WMD")
        self.recent_files_limit_var.set(5)
        self.default_location_var.set("")
        self.log_level_var.set("INFO")
        self.com_timeout_var.set(30)
        self.theme_var.set("dark")
        self.ui_density_var.set("normal")

    def apply_settings(self):
        """Apply settings and close dialog."""
        # Update settings dict with new values
        self.settings['auto_open_results'] = self.auto_open_var.get()
        self.settings['default_component'] = self.default_component_var.get()
        self.settings['recent_files_limit'] = self.recent_files_limit_var.get()
        self.settings['default_location'] = self.default_location_var.get()
        self.settings['log_level'] = self.log_level_var.get()
        self.settings['com_timeout'] = self.com_timeout_var.get()
        self.settings['theme'] = self.theme_var.get()
        self.settings['ui_density'] = self.ui_density_var.get()

        self.result = self.settings
        self.destroy()

    def cancel(self):
        """Cancel and close dialog."""
        self.result = None
        self.destroy()


class MainWindow(tk.Tk):
    """
    Main application window for the UCO to UDO Reconciliation tool.
    """
    
    def __init__(self):
        """Initialize the main window and UI components."""
        super().__init__()

        # Initialize settings
        self.settings = {
            'auto_open_results': True,
            'default_component': "WMD",
            'recent_files_limit': 5,
            'default_location': "",
            'log_level': "INFO",
            'com_timeout': 30,
            'theme': "dark",
            'ui_density': "normal",
            'recent_files': {
                'reconciliation': [],
                'trial_balance': [],
                'uco_to_udo': []
            }
        }

        # Load settings if available
        self.load_settings()

        # Setup UI
        self.title('UCO to UDO Reconciliation')
        self.geometry('800x650')
        self.minsize(650, 550)  # Set minimum window size

        # Make the window resizable
        self.resizable(True, True)

        # Status variables
        self.processing = False
        self.last_result_file = None

        # Configure root grid
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)  # Content frame

        # Apply theme
        self.apply_theme(self.settings.get('theme', 'dark'))

        # Initialize UI components
        self.create_menu()
        self.create_header()
        self.create_content()
        self.create_statusbar()

        # Set up logging
        self.logger = self.setup_logging()

        # Configure log text tags for different levels
        self.configure_log_colors()

        # Load icon
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            project_root = os.path.abspath(os.path.join(script_dir, '..', '..', '..'))
            icon_path = os.path.join(project_root, "diamond_icon.ico")
            self.iconbitmap(icon_path)
        except Exception as e:
            self.logger.warning(f"Could not load icon: {e}")

    def configure_log_colors(self) -> None:
        """Configure log text colors based on current theme."""
        theme = self.settings.get('theme', 'dark')

        if theme == 'light':
            self.log_text.tag_configure("error", foreground="#D00000")
            self.log_text.tag_configure("warning", foreground="#FF6600")
            self.log_text.tag_configure("info", foreground="#000000")
            self.log_text.tag_configure("debug", foreground="#666666")
        elif theme == 'blue':
            self.log_text.tag_configure("error", foreground="#FF0000")
            self.log_text.tag_configure("warning", foreground="#FFA500")
            self.log_text.tag_configure("info", foreground="#FFFFFF")
            self.log_text.tag_configure("debug", foreground="#88BBDD")
        else:  # dark and any other themes
            self.log_text.tag_configure("error", foreground="#FF0000")
            self.log_text.tag_configure("warning", foreground="#FFA500")
            self.log_text.tag_configure("info", foreground="#FFFFFF")
            self.log_text.tag_configure("debug", foreground="#AAAAAA")

    def apply_theme(self, theme_name: str) -> None:
        """
        Apply the selected theme to the application.

        Args:
            theme_name: The name of the theme to apply ('dark', 'light', 'system', 'blue')

        Returns:
            None
        """
        script_dir = os.path.dirname(os.path.abspath(__file__))
        project_root = os.path.abspath(os.path.join(script_dir, '..', '..', '..'))

        # Define theme styles
        style = ttk.Style()

        if theme_name == 'dark':
            # Apply Forest Dark theme
            tcl_file_path = os.path.join(project_root, 'forest-dark.tcl')
            try:
                self.tk.call('source', tcl_file_path)
                style.theme_use('forest-dark')
                self.configure(bg='#313131')

                # Configure custom styles
                style.configure("Header.TFrame", background="#1E1E1E")
                style.configure("Title.TLabel", font=("Arial", 16, "bold"), foreground="#FFFFFF", background="#1E1E1E")
                style.configure("Subtitle.TLabel", font=("Arial", 10), foreground="#BBBBBB", background="#1E1E1E")
                style.configure("StatusBar.TFrame", background="#2D2D2D")
                style.configure("StatusBar.TLabel", foreground="#AAAAAA", background="#2D2D2D", font=("TkDefaultFont", 8))
                style.configure("SectionHeader.TLabel", font=("TkDefaultFont", 10, "bold"), foreground="#FFFFFF")
                style.configure("Large.TButton", font=("TkDefaultFont", 11, "bold"))

                # Configure log colors
                if hasattr(self, 'log_text'):
                    self.log_text.config(bg='#232323', fg='#FFFFFF')

                # Theme preview styles
                style.configure("Preview.Dark.TFrame", background="#232323")
                style.configure("Preview.Light.TFrame", background="#F0F0F0")
                style.configure("Preview.System.TFrame", background="#E8E8E8")
                style.configure("Preview.Blue.TFrame", background="#1E3A5F")

            except Exception as e:
                print(f"Error applying dark theme: {e}")
                # Fall back to default theme
                style.theme_use('clam')

        elif theme_name == 'light':
            # Apply Forest Light theme
            tcl_file_path = os.path.join(project_root, 'forest-light.tcl')
            try:
                self.tk.call('source', tcl_file_path)
                style.theme_use('forest-light')
                self.configure(bg='#F0F0F0')

                # Configure custom styles
                style.configure("Header.TFrame", background="#DDDDDD")
                style.configure("Title.TLabel", font=("Arial", 16, "bold"), foreground="#333333", background="#DDDDDD")
                style.configure("Subtitle.TLabel", font=("Arial", 10), foreground="#555555", background="#DDDDDD")
                style.configure("StatusBar.TFrame", background="#E0E0E0")
                style.configure("StatusBar.TLabel", foreground="#333333", background="#E0E0E0", font=("TkDefaultFont", 8))
                style.configure("SectionHeader.TLabel", font=("TkDefaultFont", 10, "bold"), foreground="#000000")
                style.configure("Large.TButton", font=("TkDefaultFont", 11, "bold"))

                # Configure log colors
                if hasattr(self, 'log_text'):
                    self.log_text.config(bg='#FFFFFF', fg='#000000')

                # Theme preview styles
                style.configure("Preview.Dark.TFrame", background="#232323")
                style.configure("Preview.Light.TFrame", background="#F0F0F0")
                style.configure("Preview.System.TFrame", background="#E8E8E8")
                style.configure("Preview.Blue.TFrame", background="#1E3A5F")

            except Exception as e:
                print(f"Error applying light theme: {e}")
                # Fall back to default theme
                style.theme_use('default')

        elif theme_name == 'blue':
            # Create a blue theme
            try:
                # Use clam as base theme
                style.theme_use('clam')
                self.configure(bg='#1E3A5F')

                # Configure custom styles for blue theme
                style.configure(".",
                               background="#1E3A5F",
                               foreground="#FFFFFF",
                               fieldbackground="#2A4A6F")

                style.configure("TButton",
                               background="#1E3A5F",
                               foreground="#FFFFFF")

                style.map("TButton",
                         background=[('active', '#2A4A6F'), ('pressed', '#0A2A4F')],
                         foreground=[('active', '#FFFFFF')])

                style.configure("TEntry",
                               fieldbackground="#FFFFFF",
                               foreground="#000000")

                style.configure("TCombobox",
                               fieldbackground="#FFFFFF",
                               foreground="#000000")

                style.configure("Header.TFrame", background="#0A2A4F")
                style.configure("Title.TLabel", font=("Arial", 16, "bold"), foreground="#FFFFFF", background="#0A2A4F")
                style.configure("Subtitle.TLabel", font=("Arial", 10), foreground="#AADDFF", background="#0A2A4F")
                style.configure("StatusBar.TFrame", background="#0A2A4F")
                style.configure("StatusBar.TLabel", foreground="#AADDFF", background="#0A2A4F", font=("TkDefaultFont", 8))
                style.configure("SectionHeader.TLabel", font=("TkDefaultFont", 10, "bold"), foreground="#FFFFFF")
                style.configure("Large.TButton", font=("TkDefaultFont", 11, "bold"))

                # Configure log colors
                if hasattr(self, 'log_text'):
                    self.log_text.config(bg='#2A4A6F', fg='#FFFFFF')

                # Theme preview styles
                style.configure("Preview.Dark.TFrame", background="#232323")
                style.configure("Preview.Light.TFrame", background="#F0F0F0")
                style.configure("Preview.System.TFrame", background="#E8E8E8")
                style.configure("Preview.Blue.TFrame", background="#1E3A5F")

            except Exception as e:
                print(f"Error applying blue theme: {e}")
                # Fall back to default theme
                style.theme_use('clam')

        else:  # 'system' or any other value
            # Use system default theme
            try:
                style.theme_use('default')
                self.configure(bg='SystemButtonFace')

                # Configure custom styles for system theme
                bg_color = self._get_system_color('SystemButtonFace')
                fg_color = self._get_system_color('SystemButtonText')

                style.configure("Header.TFrame", background=bg_color)
                style.configure("Title.TLabel", font=("Arial", 16, "bold"), foreground=fg_color, background=bg_color)
                style.configure("Subtitle.TLabel", font=("Arial", 10), foreground=fg_color, background=bg_color)
                style.configure("StatusBar.TFrame", background=bg_color)
                style.configure("StatusBar.TLabel", foreground=fg_color, background=bg_color, font=("TkDefaultFont", 8))
                style.configure("SectionHeader.TLabel", font=("TkDefaultFont", 10, "bold"), foreground=fg_color)
                style.configure("Large.TButton", font=("TkDefaultFont", 11, "bold"))

                # Configure log colors
                if hasattr(self, 'log_text'):
                    self.log_text.config(bg='white', fg='black')

                # Theme preview styles
                style.configure("Preview.Dark.TFrame", background="#232323")
                style.configure("Preview.Light.TFrame", background="#F0F0F0")
                style.configure("Preview.System.TFrame", background="#E8E8E8")
                style.configure("Preview.Blue.TFrame", background="#1E3A5F")

            except Exception as e:
                print(f"Error applying system theme: {e}")
                # Fall back to default theme
                style.theme_use('default')

        # Apply UI density
        self.apply_ui_density()

    def _get_system_color(self, system_color: str) -> str:
        """Get a system color value."""
        try:
            return self.winfo_rgb(system_color)
        except:
            # Fallback values
            return "#E8E8E8" if system_color == "SystemButtonFace" else "#000000"

    def apply_ui_density(self) -> None:
        """Apply UI density settings."""
        density = self.settings.get('ui_density', 'normal')
        style = ttk.Style()

        if density == 'compact':
            # Smaller padding for compact mode
            style.configure("TButton", padding=2)
            style.configure("TEntry", padding=2)
            style.configure("TCombobox", padding=2)
            style.configure("TLabelframe", padding=3)
            style.configure("TNotebook", padding=2)

        elif density == 'comfortable':
            # Larger padding for comfortable mode
            style.configure("TButton", padding=8)
            style.configure("TEntry", padding=6)
            style.configure("TCombobox", padding=6)
            style.configure("TLabelframe", padding=8)
            style.configure("TNotebook", padding=6)

        else:  # 'normal'
            # Default padding
            style.configure("TButton", padding=4)
            style.configure("TEntry", padding=4)
            style.configure("TCombobox", padding=4)
            style.configure("TLabelframe", padding=5)
            style.configure("TNotebook", padding=4)
            
    def create_menu(self) -> None:
        """
        Create the application menu.
        
        Returns:
            None
        """
        menubar = tk.Menu(self)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="New Session", command=self.new_session)
        file_menu.add_separator()
        
        # Recent files submenu
        self.recent_files_menu = tk.Menu(file_menu, tearoff=0)
        file_menu.add_cascade(label="Recent Files", menu=self.recent_files_menu)
        
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.quit)
        menubar.add_cascade(label="File", menu=file_menu)
        
        # Tools menu
        tools_menu = tk.Menu(menubar, tearoff=0)
        tools_menu.add_command(label="Settings...", command=self.show_settings)
        tools_menu.add_command(label="View Logs Directory", command=self.open_logs_directory)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="View Instructions", command=self.view_instructions)
        help_menu.add_command(label="About", command=self.show_about)
        menubar.add_cascade(label="Help", menu=help_menu)
        
        self.config(menu=menubar)
        
        # Update recent files menu
        self.update_recent_files_menu()
        
    def update_recent_files_menu(self) -> None:
        """
        Update the recent files menu with the latest entries.
        
        Returns:
            None
        """
        # Clear current menu
        self.recent_files_menu.delete(0, tk.END)
        
        # Check if we have any recent files
        has_recent = False
        
        # Add recent reconciliation files
        if self.settings['recent_files']['reconciliation']:
            self.recent_files_menu.add_command(
                label="Reconciliation Files:", 
                state=tk.DISABLED
            )
            
            for path in self.settings['recent_files']['reconciliation']:
                self.recent_files_menu.add_command(
                    label=f"  {os.path.basename(path)}",
                    command=lambda p=path: self.load_recent_file('reconciliation', p)
                )
            has_recent = True
        
        # Add recent trial balance files
        if self.settings['recent_files']['trial_balance']:
            if has_recent:
                self.recent_files_menu.add_separator()
                
            self.recent_files_menu.add_command(
                label="Trial Balance Files:", 
                state=tk.DISABLED
            )
            
            for path in self.settings['recent_files']['trial_balance']:
                self.recent_files_menu.add_command(
                    label=f"  {os.path.basename(path)}",
                    command=lambda p=path: self.load_recent_file('trial_balance', p)
                )
            has_recent = True
            
        # Add recent UCO to UDO files
        if self.settings['recent_files']['uco_to_udo']:
            if has_recent:
                self.recent_files_menu.add_separator()
                
            self.recent_files_menu.add_command(
                label="UCO to UDO Files:", 
                state=tk.DISABLED
            )
            
            for path in self.settings['recent_files']['uco_to_udo']:
                self.recent_files_menu.add_command(
                    label=f"  {os.path.basename(path)}",
                    command=lambda p=path: self.load_recent_file('uco_to_udo', p)
                )
            has_recent = True
            
        # Add clear option if we have recent files
        if has_recent:
            self.recent_files_menu.add_separator()
            self.recent_files_menu.add_command(
                label="Clear Recent Files",
                command=self.clear_recent_files
            )
        else:
            self.recent_files_menu.add_command(
                label="No Recent Files",
                state=tk.DISABLED
            )
            
    def load_recent_file(self, file_type: str, path: str) -> None:
        """
        Load a recent file into the appropriate field.
        
        Args:
            file_type: Type of file ('reconciliation', 'trial_balance', or 'uco_to_udo')
            path: Path to the file
            
        Returns:
            None
        """
        if not os.path.exists(path):
            messagebox.showerror(
                "File Not Found",
                f"The file no longer exists:\n{path}\n\nIt will be removed from recent files."
            )
            
            # Remove from recent files
            if path in self.settings['recent_files'][file_type]:
                self.settings['recent_files'][file_type].remove(path)
                self.save_settings()
                self.update_recent_files_menu()
            return
                
        # Set the file in the appropriate field
        if file_type == 'reconciliation':
            self.target_file_frame.set_file_path(path)
        elif file_type == 'trial_balance':
            self.trial_balance_frame.set_file_path(path)
        elif file_type == 'uco_to_udo':
            self.uco_to_udo_frame.set_file_path(path)
            
    def clear_recent_files(self) -> None:
        """
        Clear all recent files.
        
        Returns:
            None
        """
        self.settings['recent_files'] = {
            'reconciliation': [],
            'trial_balance': [],
            'uco_to_udo': []
        }
        self.save_settings()
        self.update_recent_files_menu()
        
    def new_session(self) -> None:
        """
        Start a new session by clearing all fields.
        
        Returns:
            None
        """
        if self.processing:
            messagebox.showerror(
                "Operation in Progress",
                "Cannot start a new session while an operation is in progress."
            )
            return
            
        # Confirm with user
        if (self.target_file_frame.get_file_path() or 
            self.trial_balance_frame.get_file_path() or 
            self.uco_to_udo_frame.get_file_path()):
            if not messagebox.askyesno(
                "Confirm New Session",
                "This will clear all current selections. Continue?"
            ):
                return
                
        # Clear fields
        self.target_file_frame.set_file_path("")
        self.trial_balance_frame.set_file_path("")
        self.uco_to_udo_frame.set_file_path("")
        
        # Reset component to default
        self.component_name_combo.set(self.settings['default_component'])
        
        # Clear log
        self.log_text.delete(1.0, tk.END)
        self.logger.info("Started new session")
        
        # Reset progress
        self.progress_bar['value'] = 0
        
        # Reset status
        self.update_status("Ready")
        
    def show_settings(self) -> None:
        """
        Show the settings dialog.

        Returns:
            None
        """
        dialog = SettingsDialog(self, self.settings)

        if dialog.result:
            # Store old theme for comparison
            old_theme = self.settings.get('theme', 'dark')
            old_density = self.settings.get('ui_density', 'normal')

            # Update settings
            self.settings = dialog.result
            self.save_settings()

            # Apply new settings
            self.component_name_combo.set(self.settings['default_component'])

            # Update log level
            for handler in self.logger.handlers:
                if isinstance(handler, logging.FileHandler):
                    level = getattr(logging, self.settings['log_level'])
                    handler.setLevel(level)

            # Apply theme changes if theme or density settings changed
            current_theme = self.settings.get('theme', 'dark')
            current_density = self.settings.get('ui_density', 'normal')

            if old_theme != current_theme or old_density != current_density:
                self.apply_theme(current_theme)
                self.configure_log_colors()
                self.logger.info(f"Theme changed to {current_theme}, UI density: {current_density}")

            self.logger.info(f"Settings updated: Log level set to {self.settings['log_level']}")
            
    def load_settings(self) -> None:
        """
        Load settings from file.
        
        Returns:
            None
        """
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            project_root = os.path.abspath(os.path.join(script_dir, '..', '..', '..'))
            settings_path = os.path.join(project_root, "settings.json")
            
            if os.path.exists(settings_path):
                import json
                with open(settings_path, 'r') as f:
                    loaded_settings = json.load(f)
                    
                # Update settings with loaded values, keeping defaults for missing keys
                for key, value in loaded_settings.items():
                    self.settings[key] = value
        except Exception as e:
            print(f"Error loading settings: {e}")
            
    def save_settings(self) -> None:
        """
        Save settings to file.
        
        Returns:
            None
        """
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            project_root = os.path.abspath(os.path.join(script_dir, '..', '..', '..'))
            settings_path = os.path.join(project_root, "settings.json")
            
            import json
            with open(settings_path, 'w') as f:
                json.dump(self.settings, f, indent=2)
        except Exception as e:
            print(f"Error saving settings: {e}")
            
    def update_recent_files(self, file_type: str, file_path: str) -> None:
        """
        Update recent files list.
        
        Args:
            file_type: Type of file ('reconciliation', 'trial_balance', or 'uco_to_udo')
            file_path: Path to add to recent files
            
        Returns:
            None
        """
        # Get recent files list
        recent_files = self.settings['recent_files'].get(file_type, [])
        
        # Remove if already exists
        if file_path in recent_files:
            recent_files.remove(file_path)
            
        # Add to start of list
        recent_files.insert(0, file_path)
        
        # Limit to specified number
        limit = self.settings.get('recent_files_limit', 5)
        recent_files = recent_files[:limit]
        
        # Update settings
        self.settings['recent_files'][file_type] = recent_files
        self.save_settings()
        
        # Update menu
        self.update_recent_files_menu()

    def open_logs_directory(self) -> None:
        """
        Open the logs directory in the file explorer.
        
        Returns:
            None
        """
        log_dir = os.path.join(os.getcwd(), "logs")
        
        if not os.path.exists(log_dir):
            os.makedirs(log_dir, exist_ok=True)
            
        # Open directory using platform-specific method
        if sys.platform == 'win32':
            os.startfile(log_dir)
        elif sys.platform == 'darwin':  # macOS
            os.system(f'open "{log_dir}"')
        else:  # Linux
            os.system(f'xdg-open "{log_dir}"')
            
    def view_instructions(self) -> None:
        """
        View the instructions file.
        
        Returns:
            None
        """
        script_dir = os.path.dirname(os.path.abspath(__file__))
        project_root = os.path.abspath(os.path.join(script_dir, '..', '..', '..'))
        instructions_path = os.path.join(project_root, "INSTRUCTIONS.md")
        
        if os.path.exists(instructions_path):
            # Try to open with default application
            if sys.platform == 'win32':
                os.startfile(instructions_path)
            elif sys.platform == 'darwin':  # macOS
                os.system(f'open "{instructions_path}"')
            else:  # Linux
                os.system(f'xdg-open "{instructions_path}"')
        else:
            messagebox.showerror(
                "Instructions Not Found",
                "The instructions file (INSTRUCTIONS.md) was not found."
            )
            
    def show_about(self) -> None:
        """
        Show the about dialog.
        
        Returns:
            None
        """
        about_text = """UCO to UDO Reconciliation Tool

Version 2.0

This tool automates the reconciliation process for
comparing Unfilled Customer Orders (UCO) and 
Undelivered Orders (UDO) across multiple Excel files.

© 2025 Department of Homeland Security"""

        messagebox.showinfo("About", about_text)

    def create_header(self) -> None:
        """
        Create the application header.
        
        Returns:
            None
        """
        header_frame = ttk.Frame(self, style="Header.TFrame")
        header_frame.grid(row=0, column=0, sticky="ew", padx=0, pady=0)
        
        # Configure header grid
        header_frame.columnconfigure(0, weight=1)
        
        # Title and subtitle
        title_frame = ttk.Frame(header_frame, style="Header.TFrame")
        title_frame.grid(row=0, column=0, sticky="w", padx=20, pady=10)
        
        ttk.Label(
            title_frame, 
            text="UCO to UDO Reconciliation", 
            style="Title.TLabel"
        ).grid(row=0, column=0, sticky="w")
        
        ttk.Label(
            title_frame, 
            text="Compare and reconcile financial data across Excel spreadsheets", 
            style="Subtitle.TLabel"
        ).grid(row=1, column=0, sticky="w")
        
    def create_content(self) -> None:
        """
        Create the main content area.
        
        Returns:
            None
        """
        # Main content frame with padding
        content_frame = ttk.Frame(self, padding=10)
        content_frame.grid(row=1, column=0, sticky="nsew", padx=0, pady=0)
        
        # Configure content grid
        content_frame.columnconfigure(0, weight=1)
        content_frame.rowconfigure(2, weight=1)  # Log frame
        
        # Input section
        input_frame = ttk.LabelFrame(content_frame, text="Input Files", padding=10)
        input_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        input_frame.columnconfigure(0, weight=1)
        
        # Component selection
        component_frame = ttk.Frame(input_frame)
        component_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        component_frame.columnconfigure(1, weight=1)
        
        ttk.Label(
            component_frame, 
            text="Component:",
            style="SectionHeader.TLabel"
        ).grid(row=0, column=0, sticky="w")
        
        self.component_name_combo = ttk.Combobox(
            component_frame, 
            values=["CBP", "CG", "CIS", "CYB", "FEM", "FLE", "ICE", "MGA", "MGT", "OIG", "TSA", "SS", "ST", "WMD"], 
            state="readonly",
            width=10
        )
        self.component_name_combo.set(self.settings['default_component'])
        self.component_name_combo.grid(row=0, column=1, sticky="w", padx=(5, 0))
        
        # File input frames
        self.target_file_frame = FileInputFrame(
            input_frame,
            "UCO to UDO Reconciliation File",
            "Select the main reconciliation file where results will be stored"
        )
        self.target_file_frame.grid(row=1, column=0, sticky="ew", padx=5, pady=5)
        
        self.trial_balance_frame = FileInputFrame(
            input_frame,
            "Trial Balance File",
            "Select the trial balance Excel file containing component totals"
        )
        self.trial_balance_frame.grid(row=2, column=0, sticky="ew", padx=5, pady=5)
        
        self.uco_to_udo_frame = FileInputFrame(
            input_frame,
            "UCO to UDO TIER File",
            "Select the UCO to UDO TIER file with 'UCO to UDO' sheet"
        )
        self.uco_to_udo_frame.grid(row=3, column=0, sticky="ew", padx=5, pady=5)
        
        # Action buttons
        actions_frame = ttk.Frame(content_frame, padding=(0, 5, 0, 10))
        actions_frame.grid(row=1, column=0, sticky="ew", padx=5, pady=0)
        actions_frame.columnconfigure(0, weight=1)
        actions_frame.columnconfigure(1, weight=1)
        
        # Start button
        self.start_button = ttk.Button(
            actions_frame, 
            text="Start Reconciliation", 
            command=self.start_operations,
            style="Large.TButton"
        )
        self.start_button.grid(row=0, column=0, sticky="e", padx=5, pady=10)
        
        # Reset button
        self.reset_button = ttk.Button(
            actions_frame, 
            text="Reset", 
            command=self.new_session
        )
        self.reset_button.grid(row=0, column=1, sticky="w", padx=5, pady=10)
        
        # Result actions frame (initially hidden)
        self.result_actions_frame = ttk.Frame(content_frame)
        # Only show when we have results
        
        # Log section
        log_frame = ttk.LabelFrame(content_frame, text="Process Log", padding=5)
        log_frame.grid(row=2, column=0, sticky="nsew", padx=5, pady=5)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        # Log text with scrollbar
        self.log_text = tk.Text(
            log_frame, 
            wrap=tk.WORD, 
            bg='#232323', 
            fg='white',
            height=10  # Default height
        )
        self.log_text.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
        
        log_scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        log_scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        
        # Log filter options
        log_filter_frame = ttk.Frame(log_frame)
        log_filter_frame.grid(row=1, column=0, columnspan=2, sticky="w", padx=2, pady=(5, 2))
        
        ttk.Label(log_filter_frame, text="Show levels:").pack(side=tk.LEFT, padx=(0, 5))
        
        self.show_debug_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            log_filter_frame, 
            text="Debug", 
            variable=self.show_debug_var,
            command=self.filter_log
        ).pack(side=tk.LEFT, padx=5)
        
        self.show_info_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            log_filter_frame, 
            text="Info", 
            variable=self.show_info_var,
            command=self.filter_log
        ).pack(side=tk.LEFT, padx=5)
        
        self.show_warning_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            log_filter_frame, 
            text="Warning", 
            variable=self.show_warning_var,
            command=self.filter_log
        ).pack(side=tk.LEFT, padx=5)
        
        self.show_error_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            log_filter_frame, 
            text="Error", 
            variable=self.show_error_var,
            command=self.filter_log
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            log_filter_frame,
            text="Clear Log",
            command=self.clear_log
        ).pack(side=tk.RIGHT, padx=5)
        
        # Progress bar
        self.progress_frame = ttk.Frame(content_frame, padding=(0, 5))
        self.progress_frame.grid(row=3, column=0, sticky="ew", padx=5, pady=(0, 5))
        self.progress_frame.columnconfigure(0, weight=1)
        
        self.progress_bar = ttk.Progressbar(
            self.progress_frame, 
            orient="horizontal", 
            mode="determinate",
            length=100  # Will be stretched by grid
        )
        self.progress_bar.grid(row=0, column=0, sticky="ew", padx=2, pady=2)
        
        # Progress label
        self.progress_label = ttk.Label(self.progress_frame, text="")
        self.progress_label.grid(row=1, column=0, sticky="w", padx=5)
        
    def create_statusbar(self) -> None:
        """
        Create the status bar.
        
        Returns:
            None
        """
        status_frame = ttk.Frame(self, style="StatusBar.TFrame")
        status_frame.grid(row=2, column=0, sticky="ew", padx=0, pady=0)
        
        # Status message
        self.status_label = ttk.Label(
            status_frame, 
            text="Ready",
            style="StatusBar.TLabel"
        )
        self.status_label.pack(side=tk.LEFT, padx=10, pady=2)
        
        # Link to logs directory
        logs_link = HyperlinkLabel(
            status_frame,
            text="Open Logs Folder",
            url=""  # Will be handled by command
        )
        logs_link.bind("<Button-1>", lambda e: self.open_logs_directory())
        logs_link.pack(side=tk.RIGHT, padx=10, pady=2)
        
    def update_status(self, message: str) -> None:
        """
        Update the status bar message.
        
        Args:
            message: Status message to display
            
        Returns:
            None
        """
        self.status_label.config(text=message)
        
    def filter_log(self) -> None:
        """
        Filter log entries based on selected log levels.
        
        Returns:
            None
        """
        # Store current position
        current_pos = self.log_text.index(tk.INSERT)
        
        # Get all log text with tags
        log_content = self.log_text.get("1.0", tk.END)
        
        # Clear text widget
        self.log_text.delete("1.0", tk.END)
        
        # Reinsert content
        # This is a simplified approach. In a real implementation, 
        # you would parse the log content and filter based on level.
        # For now, we'll just update tag visibility
        self.log_text.insert(tk.END, log_content)
        
        # Show/hide based on checkboxes
        if not self.show_debug_var.get():
            self.log_text.tag_configure("debug", elide=True)
        else:
            self.log_text.tag_configure("debug", elide=False)
            
        if not self.show_info_var.get():
            self.log_text.tag_configure("info", elide=True)
        else:
            self.log_text.tag_configure("info", elide=False)
            
        if not self.show_warning_var.get():
            self.log_text.tag_configure("warning", elide=True)
        else:
            self.log_text.tag_configure("warning", elide=False)
            
        if not self.show_error_var.get():
            self.log_text.tag_configure("error", elide=True)
        else:
            self.log_text.tag_configure("error", elide=False)
            
        # Restore cursor position if possible
        try:
            self.log_text.mark_set(tk.INSERT, current_pos)
            self.log_text.see(current_pos)
        except Exception:
            self.log_text.see(tk.END)
            
    def clear_log(self) -> None:
        """
        Clear the log text widget.
        
        Returns:
            None
        """
        self.log_text.delete("1.0", tk.END)
        
    def setup_logging(self) -> logging.Logger:
        """
        Set up logging for the application.
        
        Returns:
            logging.Logger: Configured logger instance
        """
        logger = logging.getLogger("MainLogger")
        logger.setLevel(logging.DEBUG)
        
        # Remove any existing handlers
        for handler in logger.handlers[:]:
            logger.removeHandler(handler)

        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

        # File handler
        log_dir = "logs"
        os.makedirs(log_dir, exist_ok=True)
        log_filename = os.path.join(log_dir, f"UCOtoUDORecon_Log_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.txt")
        file_handler = logging.FileHandler(log_filename)
        
        # Set level from settings
        level = getattr(logging, self.settings.get('log_level', 'INFO'))
        file_handler.setLevel(level)
        
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

        # GUI Text handler
        text_handler = TextHandler(self.log_text)
        text_handler.setLevel(logging.DEBUG)  # Always capture all logs to GUI
        text_handler.setFormatter(formatter)
        logger.addHandler(text_handler)

        return logger

    def browse_file(self, entry_widget: ttk.Entry) -> None:
        """
        Open a file dialog to browse for Excel files.
        
        Args:
            entry_widget: The entry widget to update with selected file path
            
        Returns:
            None
        """
        # Determine initial directory from settings
        initial_dir = self.settings.get('default_location', '')
        if not initial_dir or not os.path.exists(initial_dir):
            initial_dir = os.getcwd()
            
        filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx")],
            initialdir=initial_dir
        )
        if filename:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, filename)
            
            # Save last directory as default
            last_dir = os.path.dirname(filename)
            if last_dir:
                self.settings['default_location'] = last_dir
                self.save_settings()

    def start_operations(self) -> None:
        """
        Start the reconciliation operations.
        
        Returns:
            None
        """
        # Check if already processing
        if self.processing:
            messagebox.showinfo(
                "Operation in Progress",
                "An operation is already in progress. Please wait for it to complete."
            )
            return
            
        component_name = self.component_name_combo.get()
        target_file = self.target_file_frame.get_file_path()
        trial_balance_file = self.trial_balance_frame.get_file_path()
        uco_to_udo_file = self.uco_to_udo_frame.get_file_path()

        # Validate all fields are filled
        if not all([component_name, target_file, trial_balance_file, uco_to_udo_file]):
            messagebox.showerror("Error", "Please select all required files and component name.")
            return
            
        # Check if files exist
        missing_files = []
        if not os.path.exists(target_file):
            missing_files.append("UCO to UDO Reconciliation File")
        if not os.path.exists(trial_balance_file):
            missing_files.append("Trial Balance File")
        if not os.path.exists(uco_to_udo_file):
            missing_files.append("UCO to UDO TIER File")
            
        if missing_files:
            messagebox.showerror(
                "Files Not Found",
                f"The following files could not be found:\n\n" + 
                "\n".join(missing_files)
            )
            return
            
        # Make sure files are Excel files
        invalid_files = []
        for file_path, file_type in [
            (target_file, "UCO to UDO Reconciliation File"),
            (trial_balance_file, "Trial Balance File"),
            (uco_to_udo_file, "UCO to UDO TIER File")
        ]:
            if not file_path.lower().endswith('.xlsx'):
                invalid_files.append(f"{file_type} (must be .xlsx)")
                
        if invalid_files:
            messagebox.showerror(
                "Invalid Files",
                f"The following files have invalid formats:\n\n" + 
                "\n".join(invalid_files)
            )
            return

        # Reset progress
        self.progress_bar['value'] = 0
        self.progress_label.config(text="")
        self.update_idletasks()
        
        # Set as processing
        self.processing = True
        self.update_status("Processing...")
        self.start_button.config(state=tk.DISABLED)
        
        # Update recent files
        self.update_recent_files('reconciliation', target_file)
        self.update_recent_files('trial_balance', trial_balance_file)
        self.update_recent_files('uco_to_udo', uco_to_udo_file)

        try:
            self.logger.info(f"Operation started with component: {component_name}")
            self.logger.info(f"UCO to UDO Reconciliation File: {os.path.basename(target_file)}")
            self.logger.info(f"Trial Balance File: {os.path.basename(trial_balance_file)}")
            self.logger.info(f"UCO to UDO TIER File: {os.path.basename(uco_to_udo_file)}")

            # Create copy of target file
            self.logger.info("Creating working copy of reconciliation file...")
            self.update_progress(5, "Creating copy of target file")
            new_target_file = create_copy_of_target_file(target_file, self.logger)
            ensure_file_handle_release(new_target_file, self.logger)

            # Copy DO TB sheet
            self.logger.info(f"Copying '{component_name} Total' sheet from Trial Balance file...")
            self.update_progress(10, "Copying Trial Balance sheet")
            if not copy_and_rename_sheet(trial_balance_file, f"{component_name} Total", new_target_file, "DO TB", self.logger, insert_index=3):
                self.logger.error(f"Failed to copy sheet '{component_name} Total'.")
                self.processing = False
                self.start_button.config(state=tk.NORMAL)
                self.update_status("Error: Failed to copy Trial Balance sheet")
                return
            ensure_file_handle_release(new_target_file, self.logger)

            # Copy DO UCO to UDO sheet
            self.logger.info("Copying 'UCO to UDO' sheet from TIER file...")
            self.update_progress(15, "Copying UCO to UDO sheet")
            if not copy_and_rename_sheet(uco_to_udo_file, "UCO to UDO", new_target_file, "DO UCO to UDO", self.logger, insert_index=4):
                self.logger.error("Failed to copy 'UCO to UDO' sheet.")
                self.processing = False
                self.start_button.config(state=tk.NORMAL)
                self.update_status("Error: Failed to copy UCO to UDO sheet")
                return
            ensure_file_handle_release(new_target_file, self.logger)
            
            # Store the result file path
            self.last_result_file = new_target_file
            
            # Perform the main operation
            self.update_progress(20, "Starting reconciliation process")
            self.after(100, lambda: self.perform_main_operation(new_target_file, component_name))

        except Exception as e:
            self.logger.error(f"Error during operation: {e}", exc_info=True)
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            self.processing = False
            self.start_button.config(state=tk.NORMAL)
            self.update_status("Error: Operation failed")

    def perform_main_operation(self, new_target_file: str, component_name: str) -> None:
        """
        Perform the main reconciliation operation.
        
        Args:
            new_target_file: Path to the target Excel file
            component_name: Selected component name
            
        Returns:
            None
        """
        try:
            self.logger.info("Running main reconciliation process...")
            
            # Run the reconciliation function
            find_table_range(
                new_target_file,
                component_name,
                self.logger,
                lambda value: self.update_progress(value, None)
            )

            # Complete
            self.progress_bar['value'] = 100
            self.update_progress(100, "Completed successfully!")
            self.update_idletasks()
            self.logger.info(f"Operations completed successfully. Output file: {os.path.basename(new_target_file)}")
            
            # Show completion message
            messagebox.showinfo(
                "Complete", 
                f"Reconciliation completed successfully!\n\n" +
                f"Output file: {os.path.basename(new_target_file)}"
            )
            
            # Update status
            self.update_status("Ready - Last operation completed successfully")
            
            # Show result file actions
            if self.settings.get('auto_open_results', True):
                self.logger.info("Auto-opening result file...")
                open_excel_file(new_target_file, self.logger)

        except Exception as e:
            self.logger.error(f"Error during reconciliation: {e}", exc_info=True)
            messagebox.showerror("Error", f"An error occurred during reconciliation: {str(e)}")
            self.update_status("Error: Reconciliation failed")
        finally:
            # Reset processing state
            self.processing = False
            self.start_button.config(state=tk.NORMAL)

    def update_progress(self, value: int, message: Optional[str] = None) -> None:
        """
        Update the progress bar and label.
        
        Args:
            value: Progress value (0-100)
            message: Optional message to display
            
        Returns:
            None
        """
        self.progress_bar['value'] = value
        
        if message is not None:
            self.progress_label.config(text=message)
            
        # Update status based on progress
        if value < 100:
            status_message = f"Processing... ({value}%)"
            if message:
                status_message += f" - {message}"
            self.update_status(status_message)
        else:
            self.update_status("Complete")
            
        self.update_idletasks()