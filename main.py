
import sys
import pandas as pd
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog, QMessageBox
from PyQt6.QtCore import Qt

class AnalysisApp(QWidget):

    def __init__(self):
        super().__init__()

        # Set up the UI
        self.setWindowTitle("NLExplorer Excel Joiner")
        self.setMinimumSize(420, 260)
        # Use targeted styles only (no global QWidget background)
        self.setStyleSheet(
            "QLabel#Title { font-size: 20px; font-weight: bold; color: #2196f3; } "
            "QLabel#Subtitle { color: #888888; font-size: 13px; } "
            "QPushButton { font-size: 15px; padding: 8px; } "
            "QPushButton#Help { background: #e1e1e1; color: #333; font-size: 13px; }"
        )

        layout = QVBoxLayout()
        layout.setSpacing(16)
        layout.setContentsMargins(32, 24, 32, 24)

        # Title
        self.title = QLabel("NLExplorer Excel Joiner")
        self.title.setObjectName("Title")
        self.title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.title)

        # Subtitle
        self.subtitle = QLabel("Reformat and summarize Neurolucida Explorer Excel output files.")
        self.subtitle.setObjectName("Subtitle")
        self.subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.subtitle)

        # Main label
        self.label = QLabel("Choose an analysis type:")
        self.label.setToolTip("Select the type of summary you want to generate from your Excel file.")
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.label)

        # Define available analysis functions
        self.analysis_methods = {
            "Marker Count Summary": self.run_marker_count_analysis,
            "Dendrite Trees Summary": self.run_dendrite_analysis,
            # Add more analyses here
        }

        # Tooltips for each analysis
        tooltips = {
            "Marker Count Summary": (
                "<b>Marker Count Summary</b><br>"
                "Summarizes marker counts by type and name across all sheets.<br>"
                "<b>Input requirements:</b> Each sheet must have columns: <i>Type, Name, Qty of Markers</i>."
            ),
            "Dendrite Trees Summary": (
                "<b>Dendrite Trees Summary</b><br>"
                "Summarizes dendrite tree metrics (length, surface, volume) per tree and per sheet.<br>"
                "<b>Input requirements:</b> Each sheet must have columns: <i>Tree, Length Total(µm), Surface Total(µm²), Volume Total(µm³)</i>."
            )
        }

        # Dynamically create buttons for each analysis type
        for analysis_name in self.analysis_methods:
            button = QPushButton(analysis_name)
            button.setToolTip(tooltips.get(analysis_name, "Run this analysis."))
            button.clicked.connect(lambda checked, name=analysis_name: self.select_file(name))
            layout.addWidget(button)

        # Add Help/Info button
        help_btn = QPushButton("Help / Info")
        help_btn.setObjectName("Help")
        help_btn.setToolTip("Show information about input requirements and usage.")
        help_btn.clicked.connect(self.show_help_dialog)
        layout.addWidget(help_btn)

        self.center_ui()
        self.setLayout(layout)

    def show_help_dialog(self):
        msg = (
            "<b>NLExplorer Excel Joiner</b><br><br>"
            "<b>How to use:</b><br>"
            "1. Click an analysis type.<br>"
            "2. Select your Excel file (.xlsx) when prompted.<br>"
            "3. The tool will process the file and save a summary output in the same folder.<br><br>"
            "<b>Input requirements:</b><br>"
            "<u>Marker Count Summary</u>:<br>"
            "&nbsp;&nbsp;- Each sheet must have columns: <i>Type, Name, Qty of Markers</i>.<br>"
            "<u>Dendrite Trees Summary</u>:<br>"
            "&nbsp;&nbsp;- Each sheet must have columns: <i>Tree, Length Total(µm), Surface Total(µm²), Volume Total(µm³)</i>.<br><br>"
            "If required columns are missing, the tool will warn you and skip those sheets."
        )
        QMessageBox.information(self, "Help / Info", msg)

    def center_ui(self):
        """Center the window on the screen"""
        screen = QApplication.primaryScreen().availableGeometry()
        size = self.geometry()
        self.move((screen.width() - size.width()) // 2, (screen.height() - size.height()) // 2)

    def select_file(self, analysis_type):
        """Open a file dialog and run the selected analysis"""
        file_path, _ = QFileDialog.getOpenFileName(self, "Select File", "", "Excel Files (*.xlsx)")

        if file_path:
            self.hide()  # Hide the main window during processing
            self.analysis_methods[analysis_type](file_path)
            self.show()  # Show the main window again after analysis
        else:
            QMessageBox.warning(self, "No File Selected", "You must select a file to proceed.")

    def run_marker_count_analysis(self, file_path: str):
        """Extract and process marker count data"""
        xls = pd.ExcelFile(file_path)
        summary_data = {}

        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            
            # Ensure required columns exist
            if {'Type', 'Name', 'Qty of Markers'}.issubset(df.columns):
                # Set 'Type' and 'Name' as a multi-index and store 'Qty of Markers'
                summary_data[sheet_name] = df.set_index(['Type', 'Name'])['Qty of Markers']
            else:
                QMessageBox.warning(self, "Data Error", f"Sheet '{sheet_name}' does not contain the required columns ('Type', 'Name', 'Qty of Markers').")

        # Combine all sheets into a single DataFrame
        summary_df = pd.DataFrame(summary_data).fillna(0).astype(int)

        # Reset index so 'Type' becomes columns
        summary_df = summary_df.T.reset_index()

        # Save to a new Excel file
        output_file = file_path.replace(".xlsx", "_Summary_Output.xlsx")
        summary_df.to_excel(output_file, index=True)

        QMessageBox.information(self, "Success", f"'Marker Count Summary' data saved to:\n{output_file}")
        self.show()  # Return to the main screen

    def run_dendrite_analysis(self, file_path : str):
        """Extract and process dendrite tree data"""
        xls = pd.ExcelFile(file_path)
        summary_data_byTree = {}
        summary_data_total = {}

        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            
            # Ensure the required columns are present
            if {'Tree', 'Length Total(µm)', 'Surface Total(µm²)', 'Volume Total(µm³)'}.issubset(df.columns):
                # Create a new DataFrame with 'Tree' as index and the relevant columns
                summary_data_byTree[sheet_name] = df.set_index('Tree')[['Length Total(µm)', 'Surface Total(µm²)', 'Volume Total(µm³)']]

                summary_data_total[sheet_name] = {
                    'Total Length (µm)': df['Length Total(µm)'].sum(),
                    'Mean Length (µm)': df['Length Total(µm)'].mean(),
                    'Total Surface (µm²)': df['Surface Total(µm²)'].sum(),
                    'Mean Surface (µm²)': df['Surface Total(µm²)'].mean(),
                    'Total Volume (µm³)': df['Volume Total(µm³)'].sum(),
                    'Mean Volume (µm³)': df['Volume Total(µm³)'].mean(),
                    'Number of Trees': len(df)
                    }
            else:
                QMessageBox.warning(self, "Data Error", f"Sheet '{sheet_name}' does not contain the required columns.")

        if not summary_data_byTree:
            QMessageBox.warning(self, "Data Error", "No valid data found in the selected file.")
            self.show()
            return

        # Combine all sheets into single DataFrames
        summary_byTree_df = pd.concat(summary_data_byTree.values(), keys=summary_data_byTree.keys(), names=['Tab Name'])
        summary_byTree_df = summary_byTree_df.reset_index()
        summary_byTree_df.rename(columns={'index': 'Tab Name'}, inplace=True)

        summary_total_df = pd.DataFrame.from_dict(summary_data_total, orient='index')
        summary_total_df.index.name = 'Tab Name'
        summary_total_df = summary_total_df.reset_index()

        

        # Save to a new Excel file
        output_file = file_path.replace(".xlsx", "_Summary_Output.xlsx")
        with pd.ExcelWriter(output_file) as writer:
            summary_total_df.to_excel(writer, sheet_name='Tab Summary', index=False)
            summary_byTree_df.to_excel(writer, sheet_name='Tree Details', index=False)

        QMessageBox.information(self, "Success", f"'Dendrite Trees Summary' data saved to:\n{output_file}")
        self.show()  # Return to the main screen

# Run the PyQt6 application
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AnalysisApp()
    window.show()
    sys.exit(app.exec())
