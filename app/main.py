# -*- coding: utf-8 -*-
"""
Created on Wed Jan 10 11:12:35 2024

@author: TOBHI

Updated on 9/11/2025 BECHY

Added a few useful buttons
Removed the Flange folder filter requirements
Set default project to first folder in the parent
Had AI remove redundancies in code
Added VBS file next to python to make it easier to run, "Home" Python hard set to C:\ProgramData\anaconda3\pythonw.exe
"""

import os
import glob
from datetime import datetime as dt
import pandas as pd
import json

from tkinter import Tk, filedialog

from PyQt5.QtWidgets import (
    QApplication, QLabel, QPushButton, QGridLayout, QWidget, QComboBox,
    QButtonGroup, QRadioButton, QMessageBox, QMenuBar, QLineEdit, QSpinBox,
    QHBoxLayout, QTreeWidget, QTreeWidgetItem, QVBoxLayout, QCheckBox,
    QProgressDialog, QHeaderView
)
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QEventLoop, Qt, QTimer

from tower_bolt_package.funcs import discover_folders, discover_xmls, find_duplicate_xmls
from tower_bolt_package.flange import Flange
from tower_bolt_package.reporting import generate_pdf, write_to_excel


# ----------------------------
# Constants and simple helpers
# ----------------------------

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

tower_patterns = [r"[a-z]{1,3}[ ]{0,1}[a-z]{0,1}[0-9]{1,3}"]
flange_patterns = []

icon_path = os.path.join("tower_bolt_package", "Vestas_Icon_BlueSky01_Service-tools_RGB.png")
template_path = os.path.join(SCRIPT_DIR, "tower_bolt_package", "report_template.xlsx")

def load_config():
    """Load configuration from tension_config.json."""
    config_path = os.path.join(SCRIPT_DIR, "tension_config.json")
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            pass
    return {"parent_path": SCRIPT_DIR}

def save_config(config):
    """Save configuration to tension_config.json."""
    config_path = os.path.join(SCRIPT_DIR, "tension_config.json")
    try:
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2)
    except IOError:
        pass


def show_msg(title: str, text: str, icon=QMessageBox.Information):
    """Topmost, modal message box with clean text."""
    msg = QMessageBox()
    msg.setIcon(icon)
    msg.setWindowTitle(title)
    msg.setText(text)
    msg.setWindowIcon(QIcon(icon_path))
    msg.setWindowFlag(Qt.WindowStaysOnTopHint, True)
    msg.setWindowModality(Qt.ApplicationModal)
    msg.exec()


def show_info(title: str, text: str):
    """Info popup."""
    show_msg(title, text, QMessageBox.Information)


def show_warn(title: str, text: str):
    """Warning popup."""
    show_msg(title, text, QMessageBox.Warning)


def has_required_xmls(flange_path: str) -> bool:
    """True if both XML groups exist in the flange folder."""
    return all((discover_xmls(flange_path, "first"),
                discover_xmls(flange_path, "second")))


def existing_report_flags(flange_path: str, project: str, tower: str, flange: str):
    """Return booleans for PDF and XLSX presence for this flange."""
    base = f"Report-{project}_{tower}_{flange}-*"
    pdf_found = any(glob.iglob(os.path.join(flange_path, base + ".pdf")))
    xlsx_found = any(glob.iglob(os.path.join(flange_path, base + ".xlsx")))
    return pdf_found, xlsx_found


def reports_exist(flange_path: str, project: str, tower: str, flange: str,
                  want_pdf: bool, want_xlsx: bool) -> bool:
    """True if reports of the chosen types already exist."""
    has_pdf, has_xlsx = existing_report_flags(flange_path, project, tower, flange)
    if want_pdf and want_xlsx:
        return has_pdf and has_xlsx
    return (want_pdf and has_pdf) or (want_xlsx and has_xlsx)


def delete_existing_reports(flange_path: str, project: str, tower: str, flange: str,
                            del_pdf: bool, del_xlsx: bool):
    """Delete existing chosen report types for this flange."""
    base = f"Report-{project}_{tower}_{flange}-*"
    if del_pdf:
        for p in glob.iglob(os.path.join(flange_path, base + ".pdf")):
            try:
                os.remove(p)
            except:
                pass
    if del_xlsx:
        for p in glob.iglob(os.path.join(flange_path, base + ".xlsx")):
            try:
                os.remove(p)
            except:
                pass


def ask_conflict_action(project: str, tower: str, flange: str,
                        has_pdf: bool, has_xlsx: bool) -> str:
    """Ask what to do when a report exists for a single flange."""
    parts = []
    if has_pdf:
        parts.append("PDF")
    if has_xlsx:
        parts.append("Excel")
    existing = " and ".join(parts) if parts else "reports"

    msg = QMessageBox()
    msg.setIcon(QMessageBox.Question)
    msg.setWindowTitle("Report already exists")
    msg.setText(f"Reports already exist for:\n{project} / {tower} / {flange}\n\n"
                f"Existing: {existing}\n\nWhat do you want to do?")
    msg.setWindowIcon(QIcon(icon_path))
    msg.setWindowFlag(Qt.WindowStaysOnTopHint, True)
    msg.setWindowModality(Qt.ApplicationModal)

    b_overwrite = msg.addButton("Overwrite", QMessageBox.AcceptRole)
    b_add = msg.addButton("Create Additional", QMessageBox.ActionRole)
    b_skip = msg.addButton("Skip", QMessageBox.RejectRole)

    msg.exec()
    if msg.clickedButton() is b_overwrite:
        return "overwrite"
    if msg.clickedButton() is b_add:
        return "additional"
    return "skip"


def latest_pdf_in_folder(folder: str, pattern: str = "Report-*.pdf"):
    """Return path to latest PDF in a folder, or None."""
    pdfs = list(glob.iglob(os.path.join(folder, pattern)))
    if not pdfs:
        return None
    pdfs.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return pdfs[0]


# ----------------------------
# Duplicate XML Finder Window
# ----------------------------

class DuplicateFinder(QWidget):
    """Tool to find and manage duplicate XML files."""
    def __init__(self, parent_path):
        super().__init__()
        self.parent_path = parent_path
        self.duplicate_groups = []
        self.setWindowTitle("Find Duplicate XML Files")
        self.setMinimumSize(900, 600)
        
        # Main layout
        layout = QVBoxLayout()
        
        # Options section
        options_layout = QHBoxLayout()
        self.checkbox_subfolders = QCheckBox("Search across all subfolders")
        self.checkbox_subfolders.setChecked(True)
        self.checkbox_subfolders.setToolTip(
            "Checked: Find duplicates across all folders\n"
            "Unchecked: Only find duplicates within the same folder"
        )
        
        # Duplicate criteria selection
        self.label_criteria = QLabel("Match by:")
        self.combo_criteria = QComboBox()
        self.combo_criteria.addItems(["File Name and Size", "File Name Only", "File Size Only"])
        self.combo_criteria.setCurrentIndex(0)
        self.combo_criteria.setToolTip(
            "Choose how to identify duplicates:\n"
            "- File Name and Size: Files must have same name AND size\n"
            "- File Name Only: Files with the same name (any size)\n"
            "- File Size Only: Files with the same size (any name)"
        )
        
        self.pushb_scan = QPushButton("Scan for Duplicates")
        self.pushb_scan.setFixedWidth(150)
        
        options_layout.addWidget(self.checkbox_subfolders)
        options_layout.addSpacing(20)
        options_layout.addWidget(self.label_criteria)
        options_layout.addWidget(self.combo_criteria)
        options_layout.addStretch()
        options_layout.addWidget(self.pushb_scan)
        
        # Results tree
        self.tree_results = QTreeWidget()
        self.tree_results.setHeaderLabels(["File Name", "Location", "Size (KB)", "Modified"])
        self.tree_results.setSelectionMode(QTreeWidget.ExtendedSelection)
        self.tree_results.setAlternatingRowColors(True)
        
        # Make columns resizable
        header = self.tree_results.header()
        header.setSectionResizeMode(0, QHeaderView.Interactive)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        
        # Info label
        self.label_info = QLabel("Click 'Scan for Duplicates' to begin")
        self.label_info.setStyleSheet("QLabel { color: #666; padding: 5px; }")
        
        # Action buttons
        buttons_layout = QHBoxLayout()
        self.pushb_delete_selected = QPushButton("Delete Selected Files")
        self.pushb_open_file = QPushButton("Open File")
        self.pushb_open_location = QPushButton("Open File Location")
        self.pushb_select_all_but_first = QPushButton("Select All But First in Each Group")
        
        self.pushb_delete_selected.setEnabled(False)
        self.pushb_open_file.setEnabled(False)
        self.pushb_open_location.setEnabled(False)
        self.pushb_select_all_but_first.setEnabled(False)
        
        buttons_layout.addWidget(self.pushb_select_all_but_first)
        buttons_layout.addWidget(self.pushb_open_file)
        buttons_layout.addWidget(self.pushb_open_location)
        buttons_layout.addWidget(self.pushb_delete_selected)
        
        # Assemble layout
        layout.addLayout(options_layout)
        layout.addWidget(QLabel("Duplicate Groups:"))
        layout.addWidget(self.tree_results)
        layout.addWidget(self.label_info)
        layout.addLayout(buttons_layout)
        
        self.setLayout(layout)
        self.setWindowIcon(QIcon(icon_path))
        
        # Connect signals
        self.pushb_scan.clicked.connect(self.scan_for_duplicates)
        self.pushb_delete_selected.clicked.connect(self.delete_selected_files)
        self.pushb_open_file.clicked.connect(self.open_file)
        self.pushb_open_location.clicked.connect(self.open_file_location)
        self.pushb_select_all_but_first.clicked.connect(self.select_all_but_first)
        self.tree_results.itemSelectionChanged.connect(self.update_button_states)
    
    def scan_for_duplicates(self):
        """Scan for duplicate XML files."""
        if not self.parent_path or not os.path.exists(self.parent_path):
            show_warn("Error", "Invalid parent path selected.")
            return
        
        # Show progress dialog
        progress = QProgressDialog("Scanning for duplicate XML files...", None, 0, 0, self)
        progress.setWindowTitle("Scanning")
        progress.setWindowModality(Qt.WindowModal)
        progress.setWindowIcon(QIcon(icon_path))
        progress.show()
        QApplication.processEvents()
        
        try:
            # Find duplicates
            search_subfolders = self.checkbox_subfolders.isChecked()
            criteria = self.combo_criteria.currentText()
            self.duplicate_groups = find_duplicate_xmls(self.parent_path, search_subfolders, criteria)
            
            # Update tree
            self.populate_tree()
            
            # Update info label
            total_duplicates = sum(len(group) for group in self.duplicate_groups)
            total_groups = len(self.duplicate_groups)
            
            if total_groups == 0:
                self.label_info.setText("No duplicate XML files found.")
                self.label_info.setStyleSheet("QLabel { color: green; padding: 5px; font-weight: bold; }")
            else:
                wasted_space = self.calculate_wasted_space()
                self.label_info.setText(
                    f"Found {total_duplicates} duplicate files in {total_groups} groups "
                    f"(~{wasted_space:.2f} MB wasted space)"
                )
                self.label_info.setStyleSheet("QLabel { color: #cc6600; padding: 5px; font-weight: bold; }")
            
        except Exception as e:
            show_warn("Error", f"Error scanning for duplicates:\n{str(e)}")
        finally:
            progress.close()
    
    def populate_tree(self):
        """Populate the tree widget with duplicate groups."""
        self.tree_results.clear()
        
        for i, group in enumerate(self.duplicate_groups, 1):
            # Create group header
            group_item = QTreeWidgetItem(self.tree_results)
            group_item.setText(0, f"Duplicate Group {i} ({len(group)} files)")
            group_item.setExpanded(True)
            
            # Style the group header
            font = group_item.font(0)
            font.setBold(True)
            group_item.setFont(0, font)
            group_item.setBackground(0, Qt.lightGray)
            group_item.setBackground(1, Qt.lightGray)
            group_item.setBackground(2, Qt.lightGray)
            group_item.setBackground(3, Qt.lightGray)
            
            # Add files in the group
            for filepath in group:
                file_item = QTreeWidgetItem(group_item)
                file_item.setData(0, Qt.UserRole, filepath)  # Store full path
                
                filename = os.path.basename(filepath)
                
                # Get relative path from parent folder
                try:
                    relative_path = os.path.relpath(filepath, self.parent_path)
                    # Show directory part of relative path (everything except filename)
                    location = os.path.dirname(relative_path)
                    if not location:
                        location = "."  # Current directory
                except:
                    # Fallback to absolute path if relative path fails
                    location = os.path.dirname(filepath)
                
                try:
                    size_kb = os.path.getsize(filepath) / 1024
                    modified_time = dt.fromtimestamp(os.path.getmtime(filepath))
                    modified_str = modified_time.strftime("%Y-%m-%d %H:%M:%S")
                except:
                    size_kb = 0
                    modified_str = "Unknown"
                
                file_item.setText(0, filename)
                file_item.setText(1, location)
                file_item.setText(2, f"{size_kb:.2f}")
                file_item.setText(3, modified_str)
        
        # Update button states
        self.update_button_states()
    
    def calculate_wasted_space(self):
        """Calculate total wasted space from duplicates in MB."""
        wasted_bytes = 0
        for group in self.duplicate_groups:
            if len(group) > 1:
                try:
                    # Size of one file times (count - 1)
                    file_size = os.path.getsize(group[0])
                    wasted_bytes += file_size * (len(group) - 1)
                except:
                    continue
        return wasted_bytes / (1024 * 1024)  # Convert to MB
    
    def select_all_but_first(self):
        """Select all duplicate files except the first one in each group."""
        self.tree_results.clearSelection()
        
        root = self.tree_results.invisibleRootItem()
        for i in range(root.childCount()):
            group_item = root.child(i)
            # Skip the first file (index 0), select the rest
            for j in range(1, group_item.childCount()):
                file_item = group_item.child(j)
                file_item.setSelected(True)
        
        self.update_button_states()
    
    def delete_selected_files(self):
        """Delete the selected files after confirmation."""
        selected_items = self.tree_results.selectedItems()
        
        # Filter to only file items (not group headers)
        files_to_delete = []
        for item in selected_items:
            filepath = item.data(0, Qt.UserRole)
            if filepath:  # Only process items with a stored filepath
                files_to_delete.append(filepath)
        
        if not files_to_delete:
            show_warn("No Selection", "No files selected for deletion.")
            return
        
        # Confirmation dialog
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Warning)
        msg.setWindowTitle("Confirm Deletion")
        msg.setText(f"Are you sure you want to delete {len(files_to_delete)} file(s)?")
        msg.setInformativeText("This action cannot be undone!")
        msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        msg.setDefaultButton(QMessageBox.No)
        msg.setWindowIcon(QIcon(icon_path))
        
        if msg.exec() == QMessageBox.Yes:
            deleted_count = 0
            failed_files = []
            
            for filepath in files_to_delete:
                try:
                    os.remove(filepath)
                    deleted_count += 1
                except Exception as e:
                    failed_files.append(f"{os.path.basename(filepath)}: {str(e)}")
            
            # Show result
            if deleted_count > 0:
                result_msg = f"Successfully deleted {deleted_count} file(s)."
                if failed_files:
                    result_msg += f"\n\nFailed to delete {len(failed_files)} file(s):\n"
                    result_msg += "\n".join(failed_files[:5])  # Show first 5 failures
                    if len(failed_files) > 5:
                        result_msg += f"\n... and {len(failed_files) - 5} more"
                
                show_info("Deletion Complete", result_msg)
                
                # Re-scan to update the list
                self.scan_for_duplicates()
            else:
                show_warn("Deletion Failed", "Failed to delete any files:\n" + "\n".join(failed_files[:10]))
    
    def open_file(self):
        """Open the selected file in the default XML viewer."""
        selected_items = self.tree_results.selectedItems()
        
        if not selected_items:
            show_warn("No Selection", "No file selected.")
            return
        
        # Get the first selected file item
        for item in selected_items:
            filepath = item.data(0, Qt.UserRole)
            if filepath and os.path.exists(filepath):
                try:
                    # Open the file with the default application
                    os.startfile(filepath)
                except Exception as e:
                    show_warn("Error Opening File", f"Could not open file:\n{str(e)}")
                break
    
    def open_file_location(self):
        """Open the folder containing the selected file."""
        selected_items = self.tree_results.selectedItems()
        
        if not selected_items:
            show_warn("No Selection", "No file selected.")
            return
        
        # Get the first selected file item
        for item in selected_items:
            filepath = item.data(0, Qt.UserRole)
            if filepath and os.path.exists(filepath):
                # Open the folder and select the file
                folder_path = os.path.dirname(filepath)
                os.startfile(folder_path)
                break
    
    def update_button_states(self):
        """Enable/disable buttons based on selection and results."""
        has_results = len(self.duplicate_groups) > 0
        has_selection = len(self.tree_results.selectedItems()) > 0
        
        # Filter selection to only file items (not group headers)
        file_items_selected = any(
            item.data(0, Qt.UserRole) is not None 
            for item in self.tree_results.selectedItems()
        )
        
        self.pushb_select_all_but_first.setEnabled(has_results)
        self.pushb_delete_selected.setEnabled(file_items_selected)
        self.pushb_open_file.setEnabled(file_items_selected)
        self.pushb_open_location.setEnabled(file_items_selected)


# ----------------------------
# Folder builder window
# ----------------------------

class FolderBuilder(QWidget):
    """Small tool to create a project tree."""
    def __init__(self):
        super().__init__()
        self.run_path = SCRIPT_DIR
        self.parent_path = SCRIPT_DIR
        self.project_name = ""
        self.tower_names = []
        self.setWindowTitle("Project Directory Builder")
        self.N_segments = 4

        # Widgets
        self.pushb_select_parent = QPushButton("Select the Project Folder Location")
        self.line_project_name = QLineEdit()
        self.line_tower_names = QLineEdit()
        self.pushb_build_directory = QPushButton("Build Directory")
        self.spinb_segments = QSpinBox()

        # Layout
        layout = QGridLayout()
        layout.addWidget(QLabel("Select Parent Folder"), 0, 0)
        layout.addWidget(QLabel("Input Project Name"), 1, 0)
        layout.addWidget(QLabel("Input Tower Names Separated by Commas"), 2, 0)
        layout.addWidget(QLabel("Input Number of Tower Segments"), 3, 0)

        layout.addWidget(self.pushb_select_parent, 0, 1, 1, 2)
        layout.addWidget(self.line_project_name, 1, 1, 1, 2)
        layout.addWidget(self.line_tower_names, 2, 1, 1, 2)
        layout.addWidget(self.spinb_segments, 3, 1, 1, 2)
        layout.addWidget(self.pushb_build_directory, 4, 0, 1, 3)
        self.setLayout(layout)

        # Spinbox rules
        self.spinb_segments.setRange(2, 10)
        self.spinb_segments.setValue(self.N_segments)

        # Wire up
        self.pushb_select_parent.clicked.connect(self.cb_pushb_select_parent)
        self.pushb_build_directory.clicked.connect(self.cb_pushb_build_directory)
        self.setWindowIcon(QIcon(icon_path))

    def cb_pushb_select_parent(self):
        """Pick parent folder."""
        try:
            Tk().withdraw()
            self.parent_path = filedialog.askdirectory(initialdir=self.parent_path)
        except Exception:
            self.parent_path = ""

    def cb_pushb_build_directory(self):
        """Create project/tower/flange folders."""
        self.project_name = self.line_project_name.text().strip()
        self.tower_names = [t for t in self.line_tower_names.text().replace(" ", "").split(",") if t]

        if not self.parent_path:
            show_info("Error", "Parent folder location not selected.")
            return
        if not self.project_name:
            show_info("Error", "No project name detected.")
            return
        if not self.tower_names:
            show_info("Error", "No tower names detected.")
            return

        project_path = os.path.join(self.parent_path, self.project_name)

        # Make flange names for N segments
        self.N_segments = int(self.spinb_segments.value())
        flange_names = []
        for i in range(self.N_segments):
            seg1 = "Base" if i == 0 else f"M{i}"
            seg2 = "Top" if i == self.N_segments - 1 else f"M{i+1}"
            flange_names.append(f"{seg1}-{seg2}")

        if not os.path.isdir(project_path):
            os.mkdir(project_path)

        for tower_name in self.tower_names:
            tower_path = os.path.join(project_path, tower_name)
            if not os.path.isdir(tower_path):
                os.mkdir(tower_path)
            for flange_name in flange_names:
                flange_path = os.path.join(tower_path, flange_name)
                if not os.path.isdir(flange_path):
                    os.mkdir(flange_path)


# ----------------------------
# Main application window
# ----------------------------

class MyWindow(QWidget):
    """Main UI and actions."""
    def __init__(self):
        super().__init__()
        self.run_path = SCRIPT_DIR
        # Load config and set parent_path
        self.config = load_config()
        self.parent_path = self.config.get("parent_path", SCRIPT_DIR)
        # Fallback to script directory if parent_path is empty
        if not self.parent_path:
            self.parent_path = SCRIPT_DIR
        self.output_location = ""
        self.setWindowTitle("Vestas Flange Reporting Tool")

        # Menus
        self.menu_bar = QMenuBar()
        menu_file = self.menu_bar.addMenu("&File")
        self.menu_file_parent = menu_file.addAction("Change Parent Folder")
        self.menu_file_build = menu_file.addAction("Build Project Folder Tree")
        menu_file.addSeparator()
        self.menu_file_duplicates = menu_file.addAction("Find Duplicate XML Files")
        menu_file.addSeparator()
        self.menu_file_reset = menu_file.addAction("Reset Options")
        self.menu_file_exit = menu_file.addAction("Exit Program")
        menu_help = self.menu_bar.addMenu("&Help")
        self.menu_help_readme = menu_help.addAction("Open README")
        self.menu_help_about = menu_help.addAction("About")
        self.pushb_refresh = QPushButton("Refresh")

        # Project row
        self.label_project = QLabel("Select a Project")
        self.combo_project = QComboBox()
        self.pushb_project = QPushButton("Run Project Reports")

        self.combo_project.addItems(discover_folders(self.parent_path, []))
        self.combo_project.setCurrentIndex(-1)
        self.pushb_project.setEnabled(False)
        self.pushb_project.setToolTip("Select a project folder.")

        # Tower row
        self.label_tower = QLabel("Select a Tower")
        self.combo_tower = QComboBox()
        self.pushb_tower = QPushButton("Run Tower Reports")
        self.pushb_open_tower_pdfs = QPushButton("Open Tower PDFs")

        self.combo_tower.setEnabled(False)
        self.pushb_tower.setEnabled(False)

        # Flange row
        self.label_flange = QLabel("Select a Flange")
        self.combo_flange = QComboBox()
        self.pushb_flange = QPushButton("Run Flange Reports")
        self.pushb_open_flange_pdf = QPushButton("Open Flange PDF")

        self.combo_flange.setEnabled(False)
        self.pushb_flange.setEnabled(False)
        self.pushb_open_flange_pdf.setEnabled(False)

        # Output options
        self.buttonGroup_format = QButtonGroup()
        self.radio_format_pdf = QRadioButton("PDF only")
        self.radio_format_excel = QRadioButton("Excel Only")
        self.radio_format_both = QRadioButton("PDF and Excel")
        for w in (self.radio_format_pdf, self.radio_format_excel, self.radio_format_both):
            self.buttonGroup_format.addButton(w)

        self.buttonGroup_location = QButtonGroup()
        self.radio_location_flange = QRadioButton("Output in Flange Folder")
        self.radio_location_select = QRadioButton("Output in Selected Folder")
        for w in (self.radio_location_flange, self.radio_location_select):
            self.buttonGroup_location.addButton(w)
        self.pushb_location_select = QPushButton("Select Output Folder")

        # Layout
        layout = QGridLayout()
        layout.addWidget(self.menu_bar, 0, 0, 1, 6)
        layout.addWidget(self.pushb_refresh, 0, 6)

        layout.addWidget(QLabel("Select Report Target"), 1, 0, 1, 4)
        layout.addWidget(QLabel("Select Output Type"), 1, 5)
        layout.addWidget(QLabel("Select Output Location"), 1, 6)

        layout.addWidget(self.label_project, 2, 0)
        layout.addWidget(self.combo_project, 2, 1, 1, 2)
        layout.addWidget(self.pushb_project, 2, 4)

        tower_btns = QWidget()
        th = QHBoxLayout(tower_btns); th.setContentsMargins(0, 0, 0, 0)
        th.addWidget(self.pushb_tower); th.addWidget(self.pushb_open_tower_pdfs)
        layout.addWidget(self.label_tower, 3, 0)
        layout.addWidget(self.combo_tower, 3, 1, 1, 2)
        layout.addWidget(tower_btns, 3, 4)

        flange_btns = QWidget()
        fh = QHBoxLayout(flange_btns); fh.setContentsMargins(0, 0, 0, 0)
        fh.addWidget(self.pushb_flange); fh.addWidget(self.pushb_open_flange_pdf)
        layout.addWidget(self.label_flange, 4, 0)
        layout.addWidget(self.combo_flange, 4, 1, 1, 2)
        layout.addWidget(flange_btns, 4, 4)

        layout.addWidget(self.radio_format_pdf, 2, 5)
        layout.addWidget(self.radio_format_excel, 3, 5)
        layout.addWidget(self.radio_format_both, 4, 5)

        layout.addWidget(self.radio_location_flange, 2, 6)
        layout.addWidget(self.radio_location_select, 3, 6)
        layout.addWidget(self.pushb_location_select, 4, 6)
        self.setLayout(layout)

        self.setWindowIcon(QIcon(icon_path))

        # Connect
        self.menu_file_parent.triggered.connect(self.cb_menu_file_parent)
        self.menu_file_build.triggered.connect(self.cb_menu_file_build)
        self.menu_file_duplicates.triggered.connect(self.cb_menu_file_duplicates)
        self.menu_file_reset.triggered.connect(self.cb_menu_file_reset)
        self.menu_file_exit.triggered.connect(self.cb_menu_file_exit)
        self.menu_help_readme.triggered.connect(self.cb_menu_help_readme)
        self.menu_help_about.triggered.connect(self.cb_menu_help_about)
        self.pushb_refresh.clicked.connect(self.cb_pushb_refresh)

        self.combo_project.currentTextChanged.connect(self.cb_select_project)
        self.combo_tower.currentTextChanged.connect(self.cb_select_tower)
        self.combo_flange.currentTextChanged.connect(self.cb_select_flange)

        self.pushb_project.clicked.connect(self.cb_run_project)
        self.pushb_tower.clicked.connect(self.cb_run_tower)
        self.pushb_flange.clicked.connect(self.cb_run_flange)

        self.pushb_open_tower_pdfs.clicked.connect(self.cb_open_tower_pdfs)
        self.pushb_open_flange_pdf.clicked.connect(self.cb_open_flange_pdf)

        self.radio_location_flange.clicked.connect(self.cb_select_output_location)
        self.radio_location_select.clicked.connect(self.cb_select_output_location)
        self.pushb_location_select.clicked.connect(self.cb_select_folder)

        # Defaults
        self.radio_location_flange.setChecked(True)
        self.pushb_location_select.setEnabled(False)
        self.radio_format_both.setChecked(True)
        QTimer.singleShot(0, self._select_first_project)

    # ---- Combo reactions ----
    def _select_first_project(self):
        # repopulate projects and trigger tower loading once, after the UI is ready
        projects = sorted(discover_folders(self.parent_path, []), key=str.lower)
        self.combo_project.blockSignals(True)
        self.combo_project.clear()
        self.combo_project.clear()
        self.combo_project.addItems(projects)
        if projects:
            self.combo_project.setCurrentIndex(0)
        self.combo_project.blockSignals(False)
        if projects:
            self.cb_select_project()

    def cb_select_project(self):
        """Load towers for the selected project."""
        # reset flange row when project changes
        self.combo_flange.clear()
        self.pushb_flange.setEnabled(False)
        self.pushb_open_flange_pdf.setEnabled(False)

        project = self.combo_project.currentText()
        project_path = os.path.join(self.parent_path, project)
        self.combo_tower.clear()
        towers = discover_folders(project_path, tower_patterns)
        self.combo_tower.addItems(towers)
        self.combo_tower.setEnabled(True)
        self.pushb_project.setEnabled(bool(towers))
        if not towers:
            self.pushb_project.setToolTip("No tower folders found.")

    def cb_select_tower(self):
        """Load flanges for the selected tower."""
        project = self.combo_project.currentText()
        tower = self.combo_tower.currentText()
        tower_path = os.path.join(self.parent_path, project, tower)
        self.combo_flange.clear()
        flanges = discover_folders(tower_path, flange_patterns)
        self.combo_flange.addItems(flanges)
        self.combo_flange.setEnabled(True)
        self.pushb_open_tower_pdfs.setEnabled(bool(flanges))
        self.pushb_tower.setEnabled(bool(flanges))
        if not flanges:
            self.pushb_tower.setToolTip("No flange folders found.")

    def cb_select_flange(self):
        """Enable flange actions and open button state."""
        project = self.combo_project.currentText()
        tower = self.combo_tower.currentText()
        flange = self.combo_flange.currentText()
        
        if not project or not tower or not flange:
            self.pushb_flange.setEnabled(False)
            self.pushb_open_flange_pdf.setEnabled(False)
            return
            
        flange_path = os.path.join(self.parent_path, project, tower, flange)
        
        # Check if path exists before checking for XMLs
        if not os.path.exists(flange_path):
            self.pushb_flange.setEnabled(False)
            self.pushb_open_flange_pdf.setEnabled(False)
            return
        
        has_xmls = has_required_xmls(flange_path)
        self.pushb_flange.setEnabled(has_xmls)
        
        # Enable open button if any PDF exists
        has_pdf = latest_pdf_in_folder(flange_path) is not None
        self.pushb_open_flange_pdf.setEnabled(has_pdf)

    # ---- Run buttons ----

    def cb_run_project(self):
        """Run all flanges in all towers. Skip when no XML or report exists."""
        if not self.run_alerts():
            return

        project = self.combo_project.currentText()
        project_path = os.path.join(self.parent_path, project)
        towers = discover_folders(project_path, tower_patterns)

        output_pdf = self.radio_format_pdf.isChecked() or self.radio_format_both.isChecked()
        output_excel = self.radio_format_excel.isChecked() or self.radio_format_both.isChecked()

        # Count total flanges for progress bar
        total_flanges = 0
        for tower in towers:
            tower_path = os.path.join(self.parent_path, project, tower)
            flanges = discover_folders(tower_path, flange_patterns)
            total_flanges += len(flanges)

        # Create progress dialog
        progress = QProgressDialog("Preparing...", "Cancel", 0, total_flanges, self)
        progress.setWindowTitle("Running Project Reports")
        progress.setWindowModality(Qt.WindowModal)
        progress.setWindowIcon(QIcon(icon_path))
        progress.setMinimumDuration(0)
        progress.setValue(0)

        skipped_no_xml = 0
        skipped_existing = 0
        exported = 0
        current = 0

        for tower in towers:
            tower_path = os.path.join(self.parent_path, project, tower)
            flanges = discover_folders(tower_path, flange_patterns)
            for flange in flanges:
                if progress.wasCanceled():
                    show_info("Cancelled", f"Report generation cancelled.\nCompleted: {exported}")
                    return

                progress.setLabelText(f"Processing: {tower} / {flange}")
                progress.setValue(current)
                QApplication.processEvents()

                flange_path = os.path.join(tower_path, flange)
                if not has_required_xmls(flange_path):
                    skipped_no_xml += 1
                    current += 1
                    continue
                if reports_exist(flange_path, project, tower, flange, output_pdf, output_excel):
                    skipped_existing += 1
                    current += 1
                    continue
                self.combo_tower.setCurrentText(tower)
                self.combo_flange.setCurrentText(flange)
                self.run_flange(ask_on_conflict=False)
                exported += 1
                current += 1

        progress.setValue(total_flanges)
        progress.close()

        show_info(
            "Flange Reports for Project Complete",
            f"Exported: {exported}\nSkipped existing: {skipped_existing}\nSkipped folders missing XML: {skipped_no_xml}"
        )

    def cb_run_tower(self):
        """Run all flanges in the selected tower. Ask once for conflicts."""
        if not self.run_alerts():
            return

        project = self.combo_project.currentText()
        tower = self.combo_tower.currentText()
        tower_path = os.path.join(self.parent_path, project, tower)
        flanges = discover_folders(tower_path, flange_patterns)

        output_pdf = self.radio_format_pdf.isChecked() or self.radio_format_both.isChecked()
        output_excel = self.radio_format_excel.isChecked() or self.radio_format_both.isChecked()

        # Pass 1: detect XML and conflicts
        have_xml = {}
        conflict = {}
        any_conflict = False
        for flange in flanges:
            fp = os.path.join(tower_path, flange)
            ok_xml = has_required_xmls(fp)
            have_xml[flange] = ok_xml
            if ok_xml:
                conflict[flange] = reports_exist(fp, project, tower, flange, output_pdf, output_excel)
                any_conflict = any_conflict or conflict[flange]
            else:
                conflict[flange] = False

        # Ask once if needed
        bulk_choice = "additional"
        if any_conflict:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Question)
            msg.setWindowTitle("Existing reports detected")
            msg.setText("Some flanges already have reports in this tower.\n\n"
                        "Choose what to do for all conflicted flanges:")
            msg.setWindowIcon(QIcon(icon_path))
            msg.setWindowFlag(Qt.WindowStaysOnTopHint, True)
            msg.setWindowModality(Qt.ApplicationModal)

            b_overwrite = msg.addButton("Overwrite all", QMessageBox.AcceptRole)
            b_add = msg.addButton("Create additional for all", QMessageBox.ActionRole)
            b_skip = msg.addButton("Skip conflicted", QMessageBox.DestructiveRole)
            b_cancel = msg.addButton("Cancel", QMessageBox.RejectRole)
            msg.exec()

            if msg.clickedButton() is b_cancel:
                return
            elif msg.clickedButton() is b_overwrite:
                bulk_choice = "overwrite"
            elif msg.clickedButton() is b_skip:
                bulk_choice = "skip"
            else:
                bulk_choice = "additional"

        # Create progress dialog
        progress = QProgressDialog("Preparing...", "Cancel", 0, len(flanges), self)
        progress.setWindowTitle(f"Running Reports for {tower}")
        progress.setWindowModality(Qt.WindowModal)
        progress.setWindowIcon(QIcon(icon_path))
        progress.setMinimumDuration(0)
        progress.setValue(0)

        # Pass 2: run and build per-line summary
        lines = []
        current = 0
        for flange in flanges:
            if progress.wasCanceled():
                show_info("Cancelled", f"Report generation cancelled.\nCompleted: {current}")
                return

            progress.setLabelText(f"Processing: {flange}")
            progress.setValue(current)
            QApplication.processEvents()

            fp = os.path.join(tower_path, flange)

            if not have_xml[flange]:
                lines.append(f"{flange}: no XML, skipped")
                current += 1
                continue

            if conflict[flange]:
                if bulk_choice == "skip":
                    lines.append(f"{flange}: conflict, skipped")
                    current += 1
                    continue
                if bulk_choice == "overwrite":
                    delete_existing_reports(fp, project, tower, flange,
                                            del_pdf=output_pdf, del_xlsx=output_excel)
                    self.combo_flange.setCurrentText(flange)
                    self.run_flange(ask_on_conflict=False)
                    lines.append(f"{flange}: overwrite, done")
                    current += 1
                    continue
                # additional
                self.combo_flange.setCurrentText(flange)
                self.run_flange(ask_on_conflict=False)
                lines.append(f"{flange}: additional, done")
                current += 1
                continue

            # No conflict
            self.combo_flange.setCurrentText(flange)
            self.run_flange(ask_on_conflict=False)
            lines.append(f"{flange}: new, done")
            current += 1

        progress.setValue(len(flanges))
        progress.close()

        show_info("Flange Reports for Tower Complete", "\n".join(lines))

    def cb_run_flange(self):
        """Run the selected flange. Ask for conflict if needed."""
        if not self.run_alerts():
            return

        project = self.combo_project.currentText()
        tower = self.combo_tower.currentText()
        flange = self.combo_flange.currentText()
        flange_path = os.path.join(self.parent_path, project, tower, flange)

        output_pdf = self.radio_format_pdf.isChecked() or self.radio_format_both.isChecked()
        output_excel = self.radio_format_excel.isChecked() or self.radio_format_both.isChecked()

        if not has_required_xmls(flange_path):
            show_warn("Flange Report", f"{flange}: no XML, skipped")
            return

        has_pdf, has_xlsx = existing_report_flags(flange_path, project, tower, flange)
        conflict = (output_pdf and has_pdf) or (output_excel and has_xlsx)

        if conflict:
            action = ask_conflict_action(project, tower, flange, has_pdf, has_xlsx)
            if action == "skip":
                show_info("Flange Report", f"{flange}: conflict, skipped")
                return
            if action == "overwrite":
                delete_existing_reports(flange_path, project, tower, flange,
                                        del_pdf=output_pdf, del_xlsx=output_excel)

        # Create progress dialog for single flange
        progress = QProgressDialog(f"Generating report for {flange}...", None, 0, 0, self)
        progress.setWindowTitle("Running Flange Report")
        progress.setWindowModality(Qt.WindowModal)
        progress.setWindowIcon(QIcon(icon_path))
        progress.setMinimumDuration(0)
        progress.show()
        QApplication.processEvents()

        # Run the report
        self.run_flange(ask_on_conflict=False)
        
        progress.close()

        # Show result based on action taken
        if conflict:
            if action == "overwrite":
                show_info("Flange Report", f"{flange}: overwrite, done")
            else:  # additional
                show_info("Flange Report", f"{flange}: additional, done")
        else:
            show_info("Flange Report", f"{flange}: new, done")

    # ---- Other actions ----

    def cb_open_tower_pdfs(self):
        """Open the latest PDF from each flange in the tower."""
        project = self.combo_project.currentText()
        tower = self.combo_tower.currentText()
        if not project or not tower:
            show_warn("Open PDFs", "Select a project and tower first.")
            return
        tower_path = os.path.join(self.parent_path, project, tower)
        opened = 0
        for flange in discover_folders(tower_path, flange_patterns):
            pdf = latest_pdf_in_folder(os.path.join(tower_path, flange))
            if pdf:
                os.startfile(pdf)
                opened += 1
        if opened == 0:
            show_warn("Open PDFs", "No PDFs found in this tower.")

    def cb_open_flange_pdf(self):
        """Open the latest PDF for the selected flange."""
        project = self.combo_project.currentText()
        tower = self.combo_tower.currentText()
        flange = self.combo_flange.currentText()
        if not project or not tower or not flange:
            show_warn("Open PDF", "Select a flange first.")
            return
        flange_path = os.path.join(self.parent_path, project, tower, flange)
        pdf = latest_pdf_in_folder(flange_path)
        if pdf:
            os.startfile(pdf)
        else:
            show_warn("Open PDF", "No PDF found for this flange.")

    def cb_select_output_location(self):
        """Switch output folder mode."""
        if self.radio_location_flange.isChecked():
            self.output_location = ""
            self.pushb_location_select.setEnabled(False)
        elif self.radio_location_select.isChecked():
            self.pushb_location_select.setEnabled(True)
            self.pushb_location_select.setToolTip("Select an output folder.")
            self.cb_select_folder()

    def cb_select_folder(self):
        """Pick a custom output folder."""
        try:
            Tk().withdraw()
            self.output_location = filedialog.askdirectory(initialdir=self.parent_path)
        except Exception:
            self.output_location = ""

    def cb_menu_file_parent(self):
        """Change parent folder for projects."""
        try:
            Tk().withdraw()
            old_path = self.parent_path
            self.parent_path = filedialog.askdirectory(initialdir=self.parent_path)

            # Save config if path changed
            if self.parent_path and self.parent_path != old_path:
                self.config["parent_path"] = self.parent_path
                save_config(self.config)

            # clear tower/flange before changing project list
            self.combo_tower.clear()
            self.combo_flange.clear()
            self.combo_tower.setEnabled(False)
            self.combo_flange.setEnabled(False)
            self.pushb_flange.setEnabled(False)
            self.pushb_open_flange_pdf.setEnabled(False)

            projects = sorted(discover_folders(self.parent_path, []), key=str.lower)
            self.combo_project.blockSignals(True)
            self.combo_project.clear()
            self.combo_project.addItems(projects)
            self.combo_project.setCurrentIndex(0 if projects else -1)
            self.combo_project.blockSignals(False)
            if projects:
                self.cb_select_project()
        except Exception:
            self.parent_path = ""

    def cb_menu_file_build(self):
        """Open the folder builder tool."""
        fb = FolderBuilder()
        fb.show()
        loop = QEventLoop()
        fb.destroyed.connect(loop.quit)
        loop.exec()

    def cb_menu_file_duplicates(self):
        """Open the duplicate XML finder tool."""
        df = DuplicateFinder(self.parent_path)
        df.show()
        loop = QEventLoop()
        df.destroyed.connect(loop.quit)
        loop.exec()

    def cb_menu_file_reset(self):
        """Reset selectors and options."""
        self.parent_path = SCRIPT_DIR
        self.output_location = ""

        # Reset config to defaults
        self.config["parent_path"] = SCRIPT_DIR
        save_config(self.config)

        self.combo_tower.clear()
        self.combo_flange.clear()
        self.combo_tower.setEnabled(False)
        self.combo_flange.setEnabled(False)
        self.pushb_flange.setEnabled(False)
        self.pushb_open_flange_pdf.setEnabled(False)

        projects = sorted(discover_folders(self.parent_path, []), key=str.lower)
        self.combo_project.blockSignals(True)
        self.combo_project.clear()
        self.combo_project.addItems(projects)
        self.combo_project.setCurrentIndex(0 if projects else -1)
        self.combo_project.blockSignals(False)
        if projects:
            self.cb_select_project()

        self.radio_format_both.setChecked(True)
        self.radio_location_flange.setChecked(True)
        self.radio_location_select.setChecked(False)

    def cb_menu_file_exit(self):
        """Close the app."""
        self.close()

    def cb_menu_help_readme(self):
        """Open README PDF."""
        os.startfile(os.path.join("tower_bolt_package", "README.pdf"))

    def cb_menu_help_about(self):
        """About dialog."""
        show_info(
            "About this tool",
            "Vestas Tower Flange Bolt Reporting Tool\n\n"
            "Automates data analysis for hydraulic smart tool flange bolts.\n\n"
            "See the README for details.\n"
            "Flange bolt technical specialist: GEOBE\n"
            "Tool developer: TOBHI"
        )

    def cb_pushb_refresh(self):
        """Refresh projects and reset options. Preserve tower selection if possible."""
        self.output_location = ""
        
        # Remember current selections
        current_project = self.combo_project.currentText()
        current_tower = self.combo_tower.currentText()

        # repopulate projects
        projects = sorted(discover_folders(self.parent_path, []), key=str.lower)
        self.combo_project.blockSignals(True)
        self.combo_project.clear()
        self.combo_project.addItems(projects)
        
        # Try to restore project selection
        if current_project and current_project in projects:
            self.combo_project.setCurrentText(current_project)
        else:
            self.combo_project.setCurrentIndex(0 if projects else -1)
        self.combo_project.blockSignals(False)
        
        if projects:
            # Reload the project (this will populate towers)
            project_path = os.path.join(self.parent_path, self.combo_project.currentText())
            towers = discover_folders(project_path, tower_patterns)
            
            self.combo_tower.blockSignals(True)
            self.combo_tower.clear()
            self.combo_tower.addItems(towers)
            self.combo_tower.setEnabled(True)
            
            # Try to restore tower selection, or default to first
            if current_tower and current_tower in towers:
                self.combo_tower.setCurrentText(current_tower)
            else:
                self.combo_tower.setCurrentIndex(0 if towers else -1)
            self.combo_tower.blockSignals(False)
            
            # Load flanges for the tower and select first
            if towers:
                tower_path = os.path.join(project_path, self.combo_tower.currentText())
                flanges = discover_folders(tower_path, flange_patterns)
                self.combo_flange.clear()
                self.combo_flange.addItems(flanges)
                self.combo_flange.setEnabled(True)
                
                # Always select first flange after refresh
                if flanges:
                    self.combo_flange.setCurrentIndex(0)
                    # Trigger the flange selection to update button states
                    self.cb_select_flange()
                else:
                    self.pushb_flange.setEnabled(False)
                    self.pushb_open_flange_pdf.setEnabled(False)
                
                # Update tower buttons
                self.pushb_open_tower_pdfs.setEnabled(bool(flanges))
                self.pushb_tower.setEnabled(bool(flanges))
            else:
                self.combo_flange.clear()
                self.combo_flange.setEnabled(False)
                self.pushb_flange.setEnabled(False)
                self.pushb_open_flange_pdf.setEnabled(False)
            
            # Update project button
            self.pushb_project.setEnabled(bool(towers))

        # defaults
        self.radio_format_both.setChecked(True)
        self.radio_location_flange.setChecked(True)
        self.radio_location_select.setChecked(False)

    # ---- Core report runner ----

    def run_flange(self, ask_on_conflict: bool = False):
        """Run report for the selected flange. Optionally prompt on conflicts."""
        # Load criteria once per call
        criteria = pd.read_excel(template_path, "Failure Criteria", index_col=0, engine='openpyxl')

        project = self.combo_project.currentText()
        tower = self.combo_tower.currentText()
        flange = self.combo_flange.currentText()
        flange_path = os.path.join(self.parent_path, project, tower, flange)

        output_pdf = self.radio_format_pdf.isChecked() or self.radio_format_both.isChecked()
        output_excel = self.radio_format_excel.isChecked() or self.radio_format_both.isChecked()

        if not has_required_xmls(flange_path):
            return

        # Optional conflict handling
        if ask_on_conflict and (output_pdf or output_excel):
            has_pdf, has_xlsx = existing_report_flags(flange_path, project, tower, flange)
            conflict = (output_pdf and has_pdf) or (output_excel and has_xlsx)
            if conflict:
                action = ask_conflict_action(project, tower, flange, has_pdf, has_xlsx)
                if action == "skip":
                    return
                if action == "overwrite":
                    delete_existing_reports(flange_path, project, tower, flange,
                                            del_pdf=output_pdf, del_xlsx=output_excel)
                # "additional" writes a new timestamped file

        # Run analysis and write outputs
        f = Flange(flange_path, dict(project=project, tower=tower, flange=flange), criteria)
        f.run()
        ts = dt.strftime(dt.today(), "%Y%m%d_%H%M%S")
        filename = f"Report-{project}_{tower}_{flange}-{ts}"

        out_dir = flange_path if self.radio_location_flange.isChecked() else self.output_location
        output_path = os.path.join(out_dir, filename)

        if self.radio_format_excel.isChecked() or self.radio_format_both.isChecked():
            write_to_excel(f, template_path, f"{output_path}.xlsx")
        if self.radio_format_pdf.isChecked() or self.radio_format_both.isChecked():
            generate_pdf(f, f"{output_path}.pdf")

    def run_alerts(self):
        """Validate required choices before running."""
        if not (self.radio_location_flange.isChecked() or self.radio_location_select.isChecked()):
            show_info("Report Alert", "No output location selected.")
            return 0
        if self.radio_location_select.isChecked() and not self.output_location:
            show_info("Report Alert", "No output location selected.")
            return 0
        if not (self.radio_format_both.isChecked() or
                self.radio_format_excel.isChecked() or
                self.radio_format_pdf.isChecked()):
            show_warn("Report Alert", "No output type selected.")
            return 0
        return 1


# ----------------------------
# App entry
# ----------------------------

if __name__ == "__main__":
    app = QApplication([])
    app.setStyle("Fusion")
    window = MyWindow()
    window.show()
    app.exec()
