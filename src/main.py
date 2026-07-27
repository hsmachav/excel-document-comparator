import PySimpleGUI as sg
import os
from pathlib import Path
from comparator import DocumentComparator
from excel_handler import ExcelHandler


class ExcelComparatorApp:
    """GUI Application for Excel Document Comparison."""

    def __init__(self):
        # Set theme
        sg.theme('DarkBlue2')
        self.window = None
        self.file1_path = None
        self.file2_path = None
        self.output_path = None

    def create_layout(self):
        """
        Create the GUI layout.
        
        Returns:
            List representing the layout
        """
        layout = [
            [sg.Text('Excel Document Comparator', font=('Arial', 18, 'bold'), pad=(10, 20))],
            
            [sg.Text('Report 1 (Earlier Date - e.g., Nov 10, 2025)', font=('Arial', 10, 'bold'))],
            [sg.Text('File:', size=(10, 1)), 
             sg.InputText(key='-FILE1-', size=(40, 1), disabled=True),
             sg.FileBrowse('Browse', file_types=(('Excel Files', '*.xlsx *.xls'),))],
            
            [sg.Text('Report 2 (Later Date - e.g., July 1, 2026)', font=('Arial', 10, 'bold'), pad=(0, (15, 0)))],
            [sg.Text('File:', size=(10, 1)), 
             sg.InputText(key='-FILE2-', size=(40, 1), disabled=True),
             sg.FileBrowse('Browse', file_types=(('Excel Files', '*.xlsx *.xls'),))],
            
            [sg.Text('Output File', font=('Arial', 10, 'bold'), pad=(0, (15, 0)))],
            [sg.Text('File:', size=(10, 1)), 
             sg.InputText(key='-OUTPUT-', size=(40, 1)),
             sg.FileSaveAs('Browse', file_types=(('Excel Files', '*.xlsx'),), default_extension='.xlsx')],
            
            [sg.Button('Compare', size=(15, 1), pad=(0, (20, 0))), 
             sg.Button('Clear', size=(15, 1), pad=(0, (20, 0))),
             sg.Button('Exit', size=(15, 1), pad=(0, (20, 0)))],
            
            [sg.Multiline(size=(80, 15), key='-LOGS-', disabled=True, pad=(0, (15, 10)))],
        ]
        return layout

    def run(self):
        """
        Run the GUI application.
        """
        self.window = sg.Window('Excel Document Comparator', self.create_layout())

        while True:
            event, values = self.window.read()

            if event == sg.WINDOW_CLOSED or event == 'Exit':
                break

            elif event == 'Compare':
                self._handle_compare(values)

            elif event == 'Clear':
                self.window['-FILE1-'].update('')
                self.window['-FILE2-'].update('')
                self.window['-OUTPUT-'].update('')
                self.window['-LOGS-'].update('')

        self.window.close()

    def _handle_compare(self, values):
        """
        Handle the compare button click.
        
        Args:
            values: Dictionary of input values from the GUI
        """
        try:
            # Get file paths
            file1 = values['-FILE1-']
            file2 = values['-FILE2-']
            output = values['-OUTPUT-']

            # Validate inputs
            if not file1 or not file2 or not output:
                self._log('ERROR: Please select both input files and specify output file.')
                return

            if not os.path.exists(file1):
                self._log(f'ERROR: File not found: {file1}')
                return

            if not os.path.exists(file2):
                self._log(f'ERROR: File not found: {file2}')
                return

            self._log('Starting comparison...')
            self._log(f'File 1: {file1}')
            self._log(f'File 2: {file2}')

            # Read Excel files
            self._log('Reading Excel files...')
            df_old = ExcelHandler.read_excel(file1)
            df_new = ExcelHandler.read_excel(file2)

            self._log(f'File 1 contains {len(df_old)} documents')
            self._log(f'File 2 contains {len(df_new)} documents')

            # Compare documents
            self._log('Comparing documents...')
            changes, summary = DocumentComparator.compare_documents(df_old, df_new)

            # Write output
            self._log('Generating output Excel file...')
            ExcelHandler.write_output_excel(output, changes, summary)

            # Display summary
            self._log('\n' + '='*60)
            self._log('COMPARISON SUMMARY')
            self._log('='*60)
            self._log(f'Total Unique Documents Changed: {summary["total_changed"]}')
            self._log(f'Total Changes Detected: {summary["total_changes"]}')
            self._log(f'  - Revision Changed Only: {summary["revision_only"]}')
            self._log(f'  - Files Changed Only: {summary["files_only"]}')
            self._log(f'  - Both Changed: {summary["both_changed"]}')
            self._log('='*60)
            self._log(f'\nOutput file created: {output}')
            self._log('\nComparison completed successfully!')

        except Exception as e:
            self._log(f'ERROR: {str(e)}')
            import traceback
            self._log(traceback.format_exc())

    def _log(self, message):
        """
        Add a message to the output log.
        
        Args:
            message: Message to log
        """
        current_text = self.window['-LOGS-'].get()
        self.window['-LOGS-'].update(current_text + message + '\n')
        self.window.refresh()


if __name__ == '__main__':
    app = ExcelComparatorApp()
    app.run()
