import sys
import pandas as pd
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog, QMessageBox

class AnalysisApp(QWidget):
    def __init__(self):
        super().__init__()

        # Set up the UI
        self.setWindowTitle("Select Analysis Type")
        self.setGeometry(0, 0, 300, 100)
        

        layout = QVBoxLayout()

        self.label = QLabel("Choose an analysis type:")
        layout.addWidget(self.label)

        # Define available analysis functions
        self.analysis_methods = {
            "Marker Count Summary": self.run_marker_count_analysis,
            "Dendrite Trees Summary": self.run_dendrite_analysis,
            # Add more analyses here
        }

        # Dynamically create buttons for each analysis type
        for analysis_name in self.analysis_methods:
            button = QPushButton(analysis_name)
            button.clicked.connect(lambda checked, name=analysis_name: self.select_file(name))
            layout.addWidget(button)
        
        self.center_ui()
        self.setLayout(layout)

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
            
            # Ensure columns exist
            if {'Name', 'Qty of Markers'}.issubset(df.columns):
                # Set 'Type' as index and store 'Qty of Markers'
                summary_data[sheet_name] = df.set_index('Name')['Qty of Markers']
            elif {'Type', 'Count'}.issubset(df.columns):
                # Set 'Type' as index and store 'Count'
                summary_data[sheet_name] = df.set_index('Type')['Count']
            else:
                QMessageBox.warning(self, "Data Error", f"Sheet '{sheet_name}' does not contain the required columns.")

        # Combine all sheets into a single DataFrame
        summary_df = pd.DataFrame(summary_data).fillna(0).astype(int)

        # Reset index so 'Type' becomes columns
        summary_df = summary_df.T.reset_index()

        # Rename columns
        summary_df.rename(columns={'index': 'Tab Name'}, inplace=True)

        # Save to a new Excel file
        output_file = file_path.replace(".xlsx", "_Summary_Output.xlsx")
        summary_df.to_excel(output_file, index=False)

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
