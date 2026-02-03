import sys
import os
import warnings
import pandas as pd
import datetime

# Suppress openpyxl print area warnings
warnings.filterwarnings('ignore', message='Print area cannot be set', category=UserWarning)
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, QFileDialog,
    QMessageBox, QListWidget, QListWidgetItem, QComboBox, QDialog, QHBoxLayout,
    QTextEdit, QProgressBar, QGraphicsDropShadowEffect, QScrollArea, QFrame
)
from PyQt5.QtGui import QFont, QPixmap, QColor, QLinearGradient, QPalette
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QPropertyAnimation, QEasingCurve

# --- Resource Path Helper for PyInstaller ---
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

# --- Constants for UI appearance and file paths ---
LOGO_PATH = "Oi360 Logo_4.png"  # Logo image file for branding
SEPARATOR_WIDTH = 2             # Separator width for UI layout

# --- Modern 2026 Theme System ---
class ThemeManager:
    """
    Manages dark and light themes with modern 2026 digital aesthetics.
    Features glassmorphism, gradient accents, and smooth color transitions.
    """
    
    DARK_THEME = {
        "name": "dark",
        "background": "qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #0f0f23, stop:0.5 #1a1a2e, stop:1 #16213e)",
        "background_solid": "#0f0f23",
        "card_bg": "rgba(30, 30, 60, 0.85)",
        "card_border": "rgba(100, 100, 255, 0.3)",
        "text_primary": "#ffffff",
        "text_secondary": "#a0a0c0",
        "accent_gradient": "qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #667eea, stop:1 #764ba2)",
        "accent_primary": "#667eea",
        "accent_secondary": "#764ba2",
        "button_bg": "qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #667eea, stop:1 #764ba2)",
        "button_hover": "qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #7b8efc, stop:1 #8b5cb6)",
        "button_pressed": "qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #5568d4, stop:1 #653d92)",
        "success_color": "#00ff88",
        "warning_color": "#ffaa00",
        "status_bg": "rgba(10, 10, 30, 0.9)",
        "progress_chunk": "qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #00ff88, stop:1 #00ccff)",
        "progress_bg": "rgba(30, 30, 60, 0.8)",
        "glow_color": "rgba(102, 126, 234, 0.5)",
        "dialog_bg": "#1a1a2e",
        "selected_color": "#00ff88"
    }
    
    LIGHT_THEME = {
        "name": "light",
        "background": "qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #f8fafc, stop:0.5 #e2e8f0, stop:1 #f1f5f9)",
        "background_solid": "#f8fafc",
        "card_bg": "rgba(255, 255, 255, 0.9)",
        "card_border": "rgba(100, 126, 234, 0.25)",
        "text_primary": "#1e293b",
        "text_secondary": "#64748b",
        "accent_gradient": "qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #3b82f6, stop:1 #8b5cf6)",
        "accent_primary": "#3b82f6",
        "accent_secondary": "#8b5cf6",
        "button_bg": "qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #3b82f6, stop:1 #8b5cf6)",
        "button_hover": "qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #4d93ff, stop:1 #9d6dff)",
        "button_pressed": "qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #2970e3, stop:1 #7a4ae0)",
        "success_color": "#22c55e",
        "warning_color": "#f59e0b",
        "status_bg": "rgba(255, 255, 255, 0.95)",
        "progress_chunk": "qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #22c55e, stop:1 #06b6d4)",
        "progress_bg": "rgba(226, 232, 240, 0.8)",
        "glow_color": "rgba(59, 130, 246, 0.4)",
        "dialog_bg": "#ffffff",
        "selected_color": "#22c55e"
    }
    
    @staticmethod
    def get_main_style(theme):
        """Returns the main window stylesheet for the given theme."""
        return f"""
            QWidget {{
                background: {theme['background']};
                color: {theme['text_primary']};
                font-family: 'Segoe UI', 'SF Pro Display', 'Roboto', sans-serif;
            }}
        """
    
    @staticmethod
    def get_button_style(theme):
        """Returns ultra-modern button stylesheet with glass effect."""
        return f"""
            QPushButton {{
                background: {theme['button_bg']};
                color: white;
                border: 1px solid {theme['card_border']};
                border-radius: 12px;
                padding: 10px 20px;
                font-weight: 600;
                font-size: 13px;
                letter-spacing: 0.5px;
            }}
            QPushButton:hover {{
                background: {theme['button_hover']};
                border: 1px solid {theme['accent_primary']};
            }}
            QPushButton:pressed {{
                background: {theme['button_pressed']};
                padding-top: 12px;
            }}
        """
    
    @staticmethod
    def get_run_button_style(theme):
        """Returns special style for the main action button."""
        return f"""
            QPushButton {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #ff6b35, stop:0.5 #f7931e, stop:1 #ffcc00);
                color: #1a1a2e;
                border: 2px solid rgba(255, 204, 0, 0.5);
                border-radius: 14px;
                padding: 12px 24px;
                font-weight: bold;
                font-size: 15px;
                letter-spacing: 1px;
            }}
            QPushButton:hover {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #ff7b45, stop:0.5 #ffa32e, stop:1 #ffdd22);
                border: 2px solid rgba(255, 220, 100, 0.7);
            }}
            QPushButton:pressed {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #e55b25, stop:0.5 #e7830e, stop:1 #e5bc00);
                padding-top: 14px;
            }}
        """
    
    @staticmethod
    def get_selected_button_style(theme):
        """Returns style for selected file buttons."""
        return f"""
            QPushButton {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 {theme['success_color']}, stop:1 #00ccff);
                color: #0f0f23;
                border: 1px solid {theme['success_color']};
                border-radius: 12px;
                padding: 10px 20px;
                font-weight: 600;
                font-size: 13px;
            }}
        """
    
    @staticmethod
    def get_progress_style(theme):
        """Returns modern progress bar stylesheet."""
        return f"""
            QProgressBar {{
                border: none;
                border-radius: 10px;
                text-align: center;
                background: {theme['progress_bg']};
                color: {theme['text_primary']};
                font-weight: bold;
                font-size: 12px;
                min-height: 24px;
            }}
            QProgressBar::chunk {{
                background: {theme['progress_chunk']};
                border-radius: 10px;
            }}
        """
    
    @staticmethod
    def get_status_box_style(theme):
        """Returns terminal-style status box stylesheet."""
        return f"""
            QTextEdit {{
                background: {theme['status_bg']};
                color: {theme['success_color']};
                border: 1px solid {theme['card_border']};
                border-radius: 12px;
                padding: 12px;
                font-family: 'Fira Code', 'JetBrains Mono', 'Consolas', monospace;
                font-size: 12px;
                line-height: 1.5;
            }}
        """
    
    @staticmethod
    def get_theme_toggle_style(theme):
        """Returns style for the theme toggle button."""
        is_dark = theme['name'] == 'dark'
        icon = "🌙" if is_dark else "☀️"
        return f"""
            QPushButton {{
                background: {theme['card_bg']};
                color: {theme['text_primary']};
                border: 2px solid {theme['card_border']};
                border-radius: 20px;
                padding: 8px 16px;
                font-size: 16px;
                min-width: 80px;
            }}
            QPushButton:hover {{
                background: {theme['accent_primary']};
                border-color: {theme['accent_primary']};
            }}
        """
    
    @staticmethod
    def get_dialog_style(theme):
        """Returns style for popup dialogs."""
        return f"""
            QDialog {{
                background: {theme['dialog_bg']};
                color: {theme['text_primary']};
            }}
            QLabel {{
                color: {theme['text_primary']};
                font-size: 13px;
            }}
            QComboBox {{
                background: {theme['card_bg']};
                color: {theme['text_primary']};
                border: 1px solid {theme['card_border']};
                border-radius: 8px;
                padding: 8px;
                min-height: 30px;
            }}
            QComboBox::drop-down {{
                border: none;
                padding-right: 10px;
            }}
            QComboBox QAbstractItemView {{
                background: {theme['dialog_bg']};
                color: {theme['text_primary']};
                selection-background-color: {theme['accent_primary']};
            }}
            QListWidget {{
                background: {theme['card_bg']};
                color: {theme['text_primary']};
                border: 1px solid {theme['card_border']};
                border-radius: 8px;
                padding: 5px;
            }}
            QListWidget::item:selected {{
                background: {theme['accent_primary']};
                color: white;
                border-radius: 4px;
            }}
            QPushButton {{
                background: {theme['button_bg']};
                color: white;
                border: none;
                border-radius: 8px;
                padding: 10px 20px;
                font-weight: 600;
            }}
            QPushButton:hover {{
                background: {theme['button_hover']};
            }}
        """

# Global theme state
current_theme = ThemeManager.DARK_THEME

# --- Utility function for logging debug messages to a file ---
def log_debug(message):
    """
    Appends debug messages with timestamps to a log file.
    Used for error tracking and debugging.
    """
    with open("debug_log.txt", "a") as f:
        f.write(f"[{datetime.datetime.now()}] {message}\n")

# --- Dialog for selecting columns from a DataFrame ---
class ColumnSelector(QDialog):
    """
    UI dialog for users to select the match column and return columns
    from a loaded Excel file. Used for both SOA and Reference files.
    """
    def __init__(self, headers, confirm_callback, is_soa=False):
        super().__init__()
        self.setWindowTitle("Select Columns")
        self.setStyleSheet(ThemeManager.get_dialog_style(current_theme))
        self.setMinimumWidth(400)
        layout = QVBoxLayout()

        match_label = QLabel("Select Match Column:")
        layout.addWidget(match_label)

        self.match_dropdown = QComboBox()
        self.match_dropdown.addItems(headers)
        layout.addWidget(self.match_dropdown)

        self.is_soa = is_soa
        if is_soa:
            layout.addWidget(QLabel("Select Date Column for Age Bucket:"))
            self.date_dropdown = QComboBox()
            self.date_dropdown.addItems(headers)
            layout.addWidget(self.date_dropdown)
            
            # Amount column selector for comparison
            layout.addWidget(QLabel("Select Amount Column (for comparison):"))
            self.amount_dropdown = QComboBox()
            self.amount_dropdown.addItem("None - No Amount Comparison")  # Optional
            self.amount_dropdown.addItems(headers)
            layout.addWidget(self.amount_dropdown)

        self.return_label = QLabel("Select Return Columns:")
        layout.addWidget(self.return_label)

        self.return_list = QListWidget()
        self.return_list.setSelectionMode(QListWidget.MultiSelection)
        for col in headers:
            item = QListWidgetItem(col)
            self.return_list.addItem(item)
        layout.addWidget(self.return_list)

        if is_soa:
            # For SOA, only match column, date column, and amount column are needed
            self.return_label.hide()
            self.return_list.hide()

        confirm_btn = QPushButton("Confirm")
        confirm_btn.clicked.connect(self.on_confirm)
        layout.addWidget(confirm_btn)

        self.setLayout(layout)
        self.confirm_callback = confirm_callback

    def on_confirm(self):
        """
        Handles confirmation and passes selected columns to callback.
        """
        match = self.match_dropdown.currentText()
        selected_returns = [i.text() for i in self.return_list.selectedItems()]
        if self.is_soa:
            # Get amount column, None if "None - No Amount Comparison" selected
            amount_col = self.amount_dropdown.currentText()
            if amount_col == "None - No Amount Comparison":
                amount_col = None
            self.confirm_callback(match, self.date_dropdown.currentText(), amount_col)
        else:
            self.confirm_callback(match, selected_returns)
        self.accept()  # Close the dialog


# --- Background worker thread for reconciliation ---
class RecoWorker(QThread):
    """
    Background thread to perform reconciliation between SOA and Reference files.
    Emits signals to update UI status and progress.
    """
    update_status = pyqtSignal(str)
    update_progress = pyqtSignal(int)
    reco_complete = pyqtSignal(pd.DataFrame)

    def __init__(self, soa_df, soa_match, soa_date_col, soa_amount_col, ref_configs):
        super().__init__()
        self.soa_df = soa_df
        self.soa_match = soa_match
        self.ref_configs = ref_configs
        self.soa_date_col = soa_date_col
        self.soa_amount_col = soa_amount_col  # User-selected amount column for comparison

    def run(self):
        df_result = self.soa_df.copy()
        
        # Store original date column values before conversion for age calculation
        original_date_col_values = None
        if self.soa_date_col and self.soa_date_col in df_result.columns:
            original_date_col_values = df_result[self.soa_date_col].copy()
            try:
                today = pd.to_datetime(datetime.datetime.today())
                # Convert to datetime for age calculation
                temp_dates = pd.to_datetime(
                    df_result[self.soa_date_col], errors='coerce', format='mixed', dayfirst=True
                )
                df_result['Age (Days)'] = (today - temp_dates).dt.days

                def bucket(days):
                    if pd.isna(days): return "Unknown"
                    elif days <= 15: return "0-15"
                    elif days <= 30: return "16-30"
                    elif days <= 60: return "31-60"
                    elif days <= 90: return "61-90"
                    elif days <= 120: return "91-120"
                    else: return "121+"

                df_result['Age Bucket'] = df_result['Age (Days)'].apply(bucket)
                df_result.insert(0, 'Age Bucket', df_result.pop('Age Bucket'))
                # Restore original date format to preserve user's date formatting
                df_result[self.soa_date_col] = original_date_col_values
            except Exception as e:
                self.update_status.emit(f"[WARNING] Age Bucket Error: {str(e)}")
                log_debug(f"Age Bucket Error: {str(e)}")
                # Restore original date values if error occurred
                if original_date_col_values is not None:
                    df_result[self.soa_date_col] = original_date_col_values
                return  # Stop here if error occurred

        # Helper function to clean invoice/match values for proper matching
        def clean_match_value(val):
            s = str(val).strip()
            # Remove leading apostrophe (Excel text marker) if present
            if s.startswith("'"):
                s = s[1:]
            # Strip leading zeros from numeric strings to normalize matching
            # e.g., '0308607218' should match '308607218'
            if s.isdigit() or (s and s.lstrip('0').isdigit()):
                s = s.lstrip('0') or '0'  # Keep at least '0' if all zeros
            return s
        
        # Create match dictionary with cleaned values
        match_sources_dict = {clean_match_value(val): [] for val in df_result[self.soa_match].astype(str).values}
        self.update_status.emit("Starting reconciliation...")

        total_steps = len(self.ref_configs) * df_result.shape[0] if df_result.shape[0] > 0 else 1
        current_step = 0

        for idx, config in enumerate(self.ref_configs):
            if config is None:
                continue
            ref_df, match_col, return_cols, _ = config
            try:
                self.update_status.emit(f"Matching Ref{idx+1} | Match = {match_col} | Returns = {', '.join(return_cols)}")
                soa_col = self.soa_match
                
                # Clean match column values: strip apostrophes and whitespace
                ref_df[match_col] = ref_df[match_col].astype(str).apply(clean_match_value)
                df_result[soa_col] = df_result[soa_col].astype(str).apply(clean_match_value)
                
                ref_extract = ref_df[[match_col] + return_cols].copy()
                ref_extract.columns = [match_col] + [f"Ref{idx+1}_{col}" for col in return_cols]
                
                df_result = pd.merge(df_result, ref_extract, left_on=soa_col, right_on=match_col, how='left')
                match_mask = df_result[f"Ref{idx+1}_{return_cols[0]}"].notna()
                for i, matched in enumerate(match_mask):
                    if matched:
                        key = clean_match_value(df_result.iloc[i][soa_col])
                        if key in match_sources_dict:
                            match_sources_dict[key].append(f"Ref{idx+1}")
                    current_step += 1
                    percent = int((current_step / total_steps) * 100)
                    self.update_progress.emit(percent)
                if idx > 0:
                    df_result.insert(df_result.shape[1], f"Separator{idx+1}", "")
            except Exception as e:
                log_debug(f"Match Error Ref{idx+1}: {str(e)}")
                self.update_status.emit(f"Error matching Ref{idx+1}: {str(e)}")
        df_result["Match Source"] = [
            ", ".join(match_sources_dict.get(clean_match_value(val), [])) if match_sources_dict.get(clean_match_value(val), []) else ""
            for val in df_result[self.soa_match].astype(str).values
        ]
        self.update_status.emit("Reconciliation Complete")
        self.update_progress.emit(100)
        if "Separator1" in df_result.columns:
            df_result = df_result.drop(columns=["Separator1"])
        
        # Clean up date columns - remove time portion (00:00:00) from date strings
        date_keywords = ['date', 'dt', 'dated']
        for col in df_result.columns:
            if any(kw in col.lower() for kw in date_keywords):
                try:
                    # Remove time portion like " 00:00:00" from date strings
                    df_result[col] = df_result[col].astype(str).str.replace(r'\s+00:00:00$', '', regex=True)
                    df_result[col] = df_result[col].replace('nan', '')
                    df_result[col] = df_result[col].replace('NaT', '')
                except Exception as e:
                    log_debug(f"Date cleanup error for column {col}: {str(e)}")
        
        filename = f"soa_reco_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        try:
            # Save result to Excel with formatted header
            with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
                df_result.to_excel(writer, index=False, sheet_name='Sheet1')
                workbook  = writer.book
                worksheet = writer.sheets['Sheet1']
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'top',
                    'fg_color': '#404040',
                    'font_color': '#FFFFFF',
                    'border': 1
                })
                for col_num, value in enumerate(df_result.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                
                # --- Amount Mismatch Highlighting ---
                # Use user-selected amount column instead of keyword detection
                all_cols = list(df_result.columns)
                
                soa_amt_col = self.soa_amount_col  # User-selected SOA amount column
                
                # Find Ref amount columns (columns starting with Ref and containing amount keywords)
                amount_keywords = ['amount', 'amt', 'value', 'total', 'sum', 'price', 'cost']
                ref_amount_cols = [c for c in all_cols 
                                   if any(kw in c.lower() for kw in amount_keywords) 
                                   and c.startswith('Ref')]
                
                if soa_amt_col and soa_amt_col in all_cols and ref_amount_cols:
                    mismatch_format = workbook.add_format({
                        'bg_color': '#FFC7CE',  # Light red
                        'font_color': '#9C0006'  # Dark red text
                    })
                    
                    soa_col_idx = all_cols.index(soa_amt_col)
                    log_debug(f"Amount Highlighting: SOA column = {soa_amt_col}, Ref columns = {ref_amount_cols}")
                    mismatch_count = 0
                    
                    # Build Amount Difference column data
                    amount_diff_data = []
                    
                    for row_idx in range(len(df_result)):
                        try:
                            soa_val = df_result.iloc[row_idx][soa_amt_col]
                            # Convert to float for comparison, handle various formats
                            if pd.notna(soa_val):
                                try:
                                    soa_num = float(str(soa_val).replace(',', '').replace('$', '').strip())
                                except (ValueError, TypeError):
                                    soa_num = None
                            else:
                                soa_num = None
                            
                            row_diffs = []
                            for ref_col in ref_amount_cols:
                                ref_val = df_result.iloc[row_idx][ref_col]
                                if pd.notna(ref_val) and soa_num is not None:
                                    try:
                                        ref_num = float(str(ref_val).replace(',', '').replace('$', '').strip())
                                        diff = soa_num - ref_num
                                        # Extract ref number from column name (e.g., "Ref1_AMOUNT" -> "Ref1")
                                        ref_name = ref_col.split('_')[0]
                                        
                                        # Format difference with sign
                                        if abs(diff) > 0.01:
                                            if diff > 0:
                                                row_diffs.append(f"{ref_name}: +{diff:.2f}")
                                            else:
                                                row_diffs.append(f"{ref_name}: {diff:.2f}")
                                            
                                            ref_col_idx = all_cols.index(ref_col)
                                            # Highlight SOA amount cell
                                            worksheet.write(row_idx + 1, soa_col_idx, soa_val, mismatch_format)
                                            # Highlight Ref amount cell
                                            worksheet.write(row_idx + 1, ref_col_idx, ref_val, mismatch_format)
                                            mismatch_count += 1
                                        else:
                                            row_diffs.append(f"{ref_name}: 0.00")
                                    except (ValueError, TypeError):
                                        pass  # Skip if value can't be converted to number
                        except Exception as e:
                            log_debug(f"Amount highlighting error at row {row_idx}: {str(e)}")
                            row_diffs = []
                        
                        amount_diff_data.append(", ".join(row_diffs) if row_diffs else "")
                    
                    # Add Amount Difference column to dataframe
                    df_result['Amount Difference'] = amount_diff_data
                    
                    # Write Amount Difference column to Excel
                    diff_col_idx = len(all_cols)  # New column at the end
                    worksheet.write(0, diff_col_idx, 'Amount Difference', header_format)
                    for row_idx, diff_val in enumerate(amount_diff_data):
                        worksheet.write(row_idx + 1, diff_col_idx, diff_val)
                    
                    self.update_status.emit(f"Amount comparison: {len(ref_amount_cols)} ref column(s) checked, {mismatch_count} mismatches highlighted")
                elif soa_amt_col:
                    log_debug(f"Amount Highlighting: No matching Ref amount columns found. SOA col = {soa_amt_col}")
                    self.update_status.emit(f"Amount comparison: No Ref amount columns detected for comparison with '{soa_amt_col}'")
                else:
                    log_debug(f"Amount Highlighting SKIPPED: No SOA amount column selected")
                    self.update_status.emit(f"Amount comparison: No amount column selected for comparison")
        except Exception as e:
            log_debug(f"Excel Write Error: {str(e)}")
            self.update_status.emit(f"Error saving Excel: {str(e)}")
            
        self.reco_complete.emit(df_result)

# --- Main application window and logic ---
class Oi360App(QWidget):
    """
    Main PyQt5 application window for the Oi360 SOA Reconciliation Tool.
    Handles UI setup, file selection, status updates, and triggers reconciliation.
    """
    def __init__(self):
        super().__init__()
        global current_theme
        self.current_theme = current_theme
        
        self.setWindowTitle("Oi360 SOA RECO TOOL")
        self.resize(850, 750)
        self.setMinimumSize(600, 400)
        
        # --- Main layout with scroll support ---
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        
        # Create scroll area
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.scroll_area.setFrameShape(QFrame.NoFrame)
        
        # Container widget inside scroll area
        self.container = QWidget()
        self.layout = QVBoxLayout(self.container)
        self.layout.setSpacing(12)
        self.layout.setContentsMargins(25, 25, 25, 25)

        # --- Top bar with logo on LEFT, title in CENTER, theme toggle on RIGHT ---
        self.top_bar = QHBoxLayout()
        
        # --- Logo placeholder (will be set in init_logo) ---
        self.logo_label = QLabel()
        self.logo_label.setFixedWidth(100)  # Reserve space for logo
        self.top_bar.addWidget(self.logo_label)
        
        self.top_bar.addStretch()
        
        # --- Header with 3D Embossed Effect (CENTERED) ---
        self.header = QLabel("Oi360 SOA RECO TOOL")
        self.header.setFont(QFont("Segoe UI", 28, QFont.Bold))
        self.header.setAlignment(Qt.AlignCenter)
        
        # Add shadow effect for 3D embossed look
        self.header_shadow = QGraphicsDropShadowEffect()
        self.header_shadow.setBlurRadius(8)
        self.header_shadow.setOffset(3, 3)
        self.header.setGraphicsEffect(self.header_shadow)
        
        self.top_bar.addWidget(self.header)
        
        self.top_bar.addStretch()
        
        # --- Theme Toggle Button (no emoji) ---
        self.theme_toggle = QPushButton("DARK MODE")
        self.theme_toggle.setFont(QFont("Segoe UI", 10, QFont.Bold))
        self.theme_toggle.setCursor(Qt.PointingHandCursor)
        self.theme_toggle.clicked.connect(self.toggle_theme)
        self.top_bar.addWidget(self.theme_toggle)
        
        self.layout.addLayout(self.top_bar)

        # --- Data Integrity Circle attribution ---
        self.attribution = QLabel("Supported by Data Integrity Circle")
        self.attribution.setFont(QFont("Segoe UI", 11))
        self.attribution.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(self.attribution)

        # --- Instructions card (no emoji) ---
        self.instructions = QLabel("""STEP-BY-STEP GUIDE:
        [1]  Select SOA file
        [2]  Select Reference files one by one (only once per session)
        [3]  Click 'Run Reconciliation'""")
        self.instructions.setFont(QFont("Segoe UI", 11))
        self.layout.addWidget(self.instructions)

        # --- Progress bar for reconciliation ---
        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.setTextVisible(True)
        self.progress.setFormat("%p%")
        self.progress.setMinimumHeight(28)
        self.layout.addWidget(self.progress)

        # --- Status box for logging messages (no emoji) ---
        self.status_box = QTextEdit()
        self.status_box.setReadOnly(True)
        self.status_box.setMinimumHeight(120)
        self.status_box.setPlaceholderText("Status messages will appear here...")
        self.layout.addWidget(self.status_box)
        
        # --- SOA file selection button (no emoji) ---
        self.soa_button = QPushButton("[+] Select SOA File")
        self.soa_button.setMinimumHeight(46)
        self.soa_button.setFont(QFont("Segoe UI", 12, QFont.Bold))
        self.soa_button.setCursor(Qt.PointingHandCursor)
        self.soa_button.clicked.connect(self.load_soa)
        self.layout.addWidget(self.soa_button)

        # --- Reference file selection buttons (no emoji) ---
        self.ref_buttons = []
        for i in range(4):
            btn = QPushButton(f"[+] Select Ref{i+1} File")
            btn.setMinimumHeight(46)
            btn.setFont(QFont("Segoe UI", 12))
            btn.setCursor(Qt.PointingHandCursor)
            btn.clicked.connect(lambda _, x=i: self.load_ref(x))
            self.layout.addWidget(btn)
            self.ref_buttons.append(btn)

        # --- Run reconciliation button (no emoji) ---
        self.run_btn = QPushButton(">>> RUN RECONCILIATION <<<")
        self.run_btn.setMinimumHeight(54)
        self.run_btn.setFont(QFont("Segoe UI", 14, QFont.Bold))
        self.run_btn.setCursor(Qt.PointingHandCursor)
        self.run_btn.clicked.connect(self.run_reco)
        self.layout.addWidget(self.run_btn)
        
        # Add stretch at end for better spacing
        self.layout.addStretch()

        # Set up scroll area
        self.scroll_area.setWidget(self.container)
        main_layout.addWidget(self.scroll_area)

        # --- Data holders for SOA and Reference files ---
        self.soa_df = None
        self.soa_match = None
        self.soa_date_col = None
        self.soa_amount_col = None  # User-selected amount column for comparison
        self.refs = [None] * 4
        self.soa_selected = False
        self.ref_selected = [False] * 4

        # Apply initial theme
        self.apply_theme()
        self.init_logo()
    
    def toggle_theme(self):
        """Toggles between dark and light themes."""
        global current_theme
        if self.current_theme['name'] == 'dark':
            self.current_theme = ThemeManager.LIGHT_THEME
            current_theme = ThemeManager.LIGHT_THEME
            self.theme_toggle.setText("LIGHT MODE")
        else:
            self.current_theme = ThemeManager.DARK_THEME
            current_theme = ThemeManager.DARK_THEME
            self.theme_toggle.setText("DARK MODE")
        self.apply_theme()
    
    def apply_theme(self):
        """Applies the current theme to all widgets."""
        theme = self.current_theme
        
        # Main window and scroll area style
        self.setStyleSheet(ThemeManager.get_main_style(theme))
        self.scroll_area.setStyleSheet(f"""
            QScrollArea {{
                background: transparent;
                border: none;
            }}
            QScrollBar:vertical {{
                background: {theme['card_bg']};
                width: 12px;
                border-radius: 6px;
            }}
            QScrollBar::handle:vertical {{
                background: {theme['accent_primary']};
                border-radius: 6px;
                min-height: 30px;
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                height: 0px;
            }}
        """)
        self.container.setStyleSheet(f"background: transparent;")
        
        # Header style with 3D embossed effect
        if theme['name'] == 'dark':
            shadow_color = QColor(0, 0, 0, 180)  # Dark shadow for dark theme
            text_color = "#ffffff"
        else:
            shadow_color = QColor(100, 100, 100, 120)  # Lighter shadow for light theme
            text_color = "#1e293b"
        
        self.header_shadow.setColor(shadow_color)
        self.header.setStyleSheet(f"""
            QLabel {{
                color: {text_color};
                background: transparent;
                padding: 5px;
            }}
        """)
        
        # Attribution style
        self.attribution.setStyleSheet(f"""
            QLabel {{
                color: {theme['text_secondary']};
                font-style: italic;
                background: transparent;
                padding: 8px;
                border-radius: 8px;
            }}
        """)
        
        # Instructions style with card effect
        self.instructions.setStyleSheet(f"""
            QLabel {{
                color: {theme['text_secondary']};
                background: {theme['card_bg']};
                border: 1px solid {theme['card_border']};
                border-radius: 12px;
                padding: 15px;
            }}
        """)
        
        # Theme toggle button
        self.theme_toggle.setStyleSheet(ThemeManager.get_theme_toggle_style(theme))
        
        # Progress bar
        progress_style = ThemeManager.get_progress_style(theme)
        self.progress.setStyleSheet(progress_style)
        
        # Status box
        self.status_box.setStyleSheet(ThemeManager.get_status_box_style(theme))
        
        # Regular buttons
        button_style = ThemeManager.get_button_style(theme)
        selected_style = ThemeManager.get_selected_button_style(theme)
        
        # SOA button
        if self.soa_selected:
            self.soa_button.setStyleSheet(selected_style)
        else:
            self.soa_button.setStyleSheet(button_style)
        
        # Reference buttons
        for i, btn in enumerate(self.ref_buttons):
            if self.ref_selected[i]:
                btn.setStyleSheet(selected_style)
            else:
                btn.setStyleSheet(button_style)
        
        # Run button with special style
        self.run_btn.setStyleSheet(ThemeManager.get_run_button_style(theme))

    def init_logo(self):
        """
        Adds the Oi360 logo to the LEFT side of the top bar.
        """
        logo_path = resource_path(LOGO_PATH)
        if os.path.exists(logo_path):
            pixmap = QPixmap(logo_path)
            # Scale logo to a decent visible size
            self.logo_label.setPixmap(pixmap.scaledToHeight(80, Qt.SmoothTransformation))
            self.logo_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            self.logo_label.setFixedWidth(100)  # Width to fit larger logo

    def load_soa(self):
        """
        Loads the SOA Excel file and prompts user to select the match column.
        """
        file_path, _ = QFileDialog.getOpenFileName(self, "Select SOA File", "", "Excel Files (*.xlsx)")
        if not file_path:
            return
        try:
            df = pd.read_excel(file_path, dtype=str)
            self.soa_df = df
            selector = ColumnSelector(df.columns.tolist(), self.save_soa_config, is_soa=True)
            selector.exec_()
            self.log_status(f"[OK] Loaded SOA file: {os.path.basename(file_path)} with {df.shape[0]} rows")
            self.log_status(f"[->] Selected Match: {self.soa_match}")
            # Mark as selected and apply theme-aware styling
            self.soa_selected = True
            self.soa_button.setStyleSheet(ThemeManager.get_selected_button_style(self.current_theme))
        except Exception as e:
            log_debug(str(e))
            QMessageBox.critical(self, "Error", str(e))

    def save_soa_config(self, match_col, date_col, amount_col):
        """
        Saves the selected match column and amount column for SOA file.
        """
        self.soa_match = match_col
        self.soa_date_col = date_col
        self.soa_amount_col = amount_col

    def load_ref(self, idx):
        """
        Loads a reference Excel file and prompts user to select match and return columns.
        """
        file_path, _ = QFileDialog.getOpenFileName(self, f"Select Ref{idx+1} File", "", "Excel Files (*.xlsx)")
        if not file_path:
            return
        try:
            df = pd.read_excel(file_path, dtype=str)
            selector = ColumnSelector(df.columns.tolist(), lambda m, r: self.save_ref_config(idx, df, m, r))
            selector.exec_()
            # Mark as selected and apply theme-aware styling
            self.ref_selected[idx] = True
            self.ref_buttons[idx].setStyleSheet(ThemeManager.get_selected_button_style(self.current_theme))
            self.log_status(f"[OK] Loaded Ref{idx+1}: {os.path.basename(file_path)} with {df.shape[0]} rows")
            self.log_status(f"[->] Selected Match: {self.refs[idx][1]} | Return: {', '.join(self.refs[idx][2])}")
        except Exception as e:
            log_debug(str(e))
            QMessageBox.critical(self, "Error", str(e))

    def save_ref_config(self, idx, df, match, returns):
        """
        Saves the selected match and return columns for a reference file.
        """
        self.refs[idx] = (df, match, returns, os.path.basename("Ref File"))

    def log_status(self, message):
        """
        Appends a message to the status box and logs it to the debug file.
        """
        self.status_box.append(message)
        log_debug(message)

    def run_reco(self):
        """
        Starts the reconciliation process in a background thread.
        """
        if self.soa_df is None or self.soa_match is None:
            QMessageBox.warning(self, "Missing Info", "Load SOA file and select match column first.")
            return

        self.progress.setValue(0)
        ref_configs = []
        for ref in self.refs:
            if ref is None:
                ref_configs.append(None)
            else:
                df, match, ret, lbl = ref
                df[match] = df[match].astype(str)
                ref_configs.append((df, match, ret, lbl))
        self.worker = RecoWorker(self.soa_df, self.soa_match, self.soa_date_col, self.soa_amount_col, ref_configs)
        self.worker.update_status.connect(self.log_status)
        self.worker.update_progress.connect(self.progress.setValue)
        self.worker.reco_complete.connect(self.save_output)
        self.worker.start()
        self.worker.start()

    def save_output(self, df):
        """
        Prompts user to save the reconciled DataFrame to an Excel file.
        """
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Reconciled File",
            f"soa_reco_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "Excel Files (*.xlsx)",
            options=options
        )
        if save_path:
            df.to_excel(save_path, index=False)
            self.log_status(f"Saved result to {save_path}")
            QMessageBox.information(self, "Done", f"Reconciliation saved as:\n{save_path}")
        else:
            self.log_status("Save cancelled.")
            QMessageBox.information(self, "Cancelled", "Save operation was cancelled.")

# --- Entry point for launching the application ---
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = Oi360App()
    window.show()
    sys.exit(app.exec_())
