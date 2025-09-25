#!/usr/bin/env python3
"""
LDCC1 Data Processor
====================

A comprehensive script for processing client cash management data with GUI interface.
Features:
- CSV file selection with file browser
- Payments processing checkbox option
- Automated workflow up to EQ online stage
- Data validation and error handling
- Professional logging and reporting

Author: LDCC1 Automation Team
Version: 1.0.0
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import os
import sys
import logging
import traceback
from datetime import datetime, date
from pathlib import Path
import json

class LDCC1Processor:
    """Main class for LDCC1 data processing application."""
    
    def __init__(self):
        """Initialize the application."""
        self.setup_logging()
        self.root = tk.Tk()
        self.csv_file_path = None
        self.process_payments = tk.BooleanVar()
        self.data = None
        self.setup_gui()
        
    def setup_logging(self):
        """Setup logging configuration."""
        log_dir = Path("logs")
        log_dir.mkdir(exist_ok=True)
        
        log_filename = f"ldcc1_processor_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        log_path = log_dir / log_filename
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_path),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
        self.logger.info("LDCC1 Processor initialized")
        
    def setup_gui(self):
        """Setup the graphical user interface."""
        self.root.title("LDCC1 Data Processor v1.0.0")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # Configure styles
        style = ttk.Style()
        style.theme_use('clam')
        
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(6, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="LDCC1 Client Cash Management Processor", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # CSV File Selection
        ttk.Label(main_frame, text="CSV File:", font=('Arial', 10, 'bold')).grid(
            row=1, column=0, sticky=tk.W, pady=5)
        
        self.file_var = tk.StringVar()
        self.file_entry = ttk.Entry(main_frame, textvariable=self.file_var, width=60)
        self.file_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(10, 5), pady=5)
        
        self.browse_button = ttk.Button(main_frame, text="Browse...", command=self.browse_file)
        self.browse_button.grid(row=1, column=2, padx=5, pady=5)
        
        # Payments Checkbox
        self.payment_checkbox = ttk.Checkbutton(
            main_frame, 
            text="Process Payments (will stop before EQ online)", 
            variable=self.process_payments,
            font=('Arial', 10, 'bold')
        )
        self.payment_checkbox.grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=10)
        
        # Process Button
        self.process_button = ttk.Button(
            main_frame, 
            text="Start Processing", 
            command=self.start_processing,
            style='Accent.TButton'
        )
        self.process_button.grid(row=3, column=1, pady=20, sticky=tk.EW)
        
        # Progress Bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            main_frame, 
            variable=self.progress_var, 
            maximum=100, 
            length=400
        )
        self.progress_bar.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        # Status Label
        self.status_var = tk.StringVar()
        self.status_var.set("Ready to process data")
        self.status_label = ttk.Label(main_frame, textvariable=self.status_var)
        self.status_label.grid(row=5, column=0, columnspan=3, pady=5)
        
        # Log Output
        ttk.Label(main_frame, text="Processing Log:", font=('Arial', 10, 'bold')).grid(
            row=6, column=0, sticky=(tk.W, tk.N), pady=(10, 5))
        
        self.log_text = scrolledtext.ScrolledText(
            main_frame, 
            height=15, 
            width=80,
            wrap=tk.WORD,
            font=('Consolas', 9)
        )
        self.log_text.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=8, column=0, columnspan=3, pady=20)
        
        self.clear_log_button = ttk.Button(button_frame, text="Clear Log", command=self.clear_log)
        self.clear_log_button.pack(side=tk.LEFT, padx=5)
        
        self.save_log_button = ttk.Button(button_frame, text="Save Log", command=self.save_log)
        self.save_log_button.pack(side=tk.LEFT, padx=5)
        
        self.exit_button = ttk.Button(button_frame, text="Exit", command=self.root.quit)
        self.exit_button.pack(side=tk.RIGHT, padx=5)
        
        # Redirect logging to GUI
        self.setup_gui_logging()
        
    def setup_gui_logging(self):
        """Setup logging to display in GUI."""
        class GUILogHandler(logging.Handler):
            def __init__(self, text_widget):
                super().__init__()
                self.text_widget = text_widget
                
            def emit(self, record):
                msg = self.format(record)
                self.text_widget.insert(tk.END, msg + '\n')
                self.text_widget.see(tk.END)
                self.text_widget.update()
                
        gui_handler = GUILogHandler(self.log_text)
        gui_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        self.logger.addHandler(gui_handler)
        
    def browse_file(self):
        """Open file browser for CSV selection."""
        file_path = filedialog.askopenfilename(
            title="Select CSV File",
            filetypes=[
                ("CSV files", "*.csv"),
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*")
            ],
            initialdir=os.getcwd()
        )
        
        if file_path:
            self.csv_file_path = file_path
            self.file_var.set(file_path)
            self.logger.info(f"Selected file: {file_path}")
            
    def validate_input(self):
        """Validate user inputs before processing."""
        if not self.csv_file_path:
            messagebox.showerror("Error", "Please select a CSV file to process")
            return False
            
        if not os.path.exists(self.csv_file_path):
            messagebox.showerror("Error", "Selected file does not exist")
            return False
            
        return True
        
    def update_progress(self, value, status="Processing..."):
        """Update progress bar and status."""
        self.progress_var.set(value)
        self.status_var.set(status)
        self.root.update()
        
    def load_data(self):
        """Load data from selected file."""
        try:
            self.update_progress(10, "Loading data file...")
            
            file_ext = Path(self.csv_file_path).suffix.lower()
            
            if file_ext == '.csv':
                self.data = pd.read_csv(self.csv_file_path)
            elif file_ext in ['.xlsx', '.xls']:
                self.data = pd.read_excel(self.csv_file_path)
            else:
                raise ValueError(f"Unsupported file format: {file_ext}")
                
            self.logger.info(f"Successfully loaded data: {len(self.data)} rows, {len(self.data.columns)} columns")
            self.logger.info(f"Columns: {list(self.data.columns)}")
            
            return True
            
        except Exception as e:
            self.logger.error(f"Error loading data: {str(e)}")
            messagebox.showerror("Data Loading Error", f"Failed to load data:\n{str(e)}")
            return False
            
    def validate_data_structure(self):
        """Validate the structure of loaded data."""
        self.update_progress(20, "Validating data structure...")
        
        try:
            # Check if data is not empty
            if self.data.empty:
                raise ValueError("The selected file is empty")
                
            # Log basic data info
            self.logger.info(f"Data shape: {self.data.shape}")
            self.logger.info(f"Data types:\n{self.data.dtypes}")
            
            # Check for common required columns (adjust as needed)
            potential_columns = ['client', 'amount', 'date', 'reference', 'balance', 'payment']
            found_columns = []
            
            for col in self.data.columns:
                col_lower = col.lower()
                for potential in potential_columns:
                    if potential in col_lower:
                        found_columns.append(col)
                        break
                        
            self.logger.info(f"Identified potential data columns: {found_columns}")
            
            return True
            
        except Exception as e:
            self.logger.error(f"Data validation error: {str(e)}")
            messagebox.showerror("Data Validation Error", f"Data validation failed:\n{str(e)}")
            return False
            
    def process_benefits(self):
        """Process benefits data."""
        self.update_progress(30, "Processing benefits...")
        
        try:
            # Example benefits processing logic
            # This would be customized based on actual data structure
            
            self.logger.info("Starting benefits processing...")
            
            # Simulate processing steps
            import time
            time.sleep(1)  # Simulate processing time
            
            # Add any benefits-specific calculations here
            benefits_total = 0
            if 'amount' in [col.lower() for col in self.data.columns]:
                amount_col = next((col for col in self.data.columns if 'amount' in col.lower()), None)
                if amount_col:
                    benefits_total = self.data[amount_col].sum()
                    
            self.logger.info(f"Benefits processing completed. Total amount: £{benefits_total:,.2f}")
            
            return True
            
        except Exception as e:
            self.logger.error(f"Benefits processing error: {str(e)}")
            return False
            
    def process_reconciliation(self):
        """Process reconciliation data."""
        self.update_progress(50, "Processing reconciliation...")
        
        try:
            self.logger.info("Starting reconciliation processing...")
            
            # Example reconciliation logic
            # This would include balance checks, validation, etc.
            
            import time
            time.sleep(1)  # Simulate processing time
            
            # Add reconciliation-specific logic here
            self.logger.info("Reconciliation processing completed successfully")
            
            return True
            
        except Exception as e:
            self.logger.error(f"Reconciliation processing error: {str(e)}")
            return False
            
    def prepare_payment_data(self):
        """Prepare payment data for EQ online."""
        self.update_progress(70, "Preparing payment data...")
        
        try:
            self.logger.info("Preparing payment data for EQ online...")
            
            # Example payment preparation logic
            # This would format data for EQ banking system
            
            import time
            time.sleep(1)  # Simulate processing time
            
            # Create output directory for payment files
            output_dir = Path("payment_output")
            output_dir.mkdir(exist_ok=True)
            
            # Generate payment summary
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            payment_summary_file = output_dir / f"payment_summary_{timestamp}.csv"
            
            # Example: Save processed data
            if self.data is not None:
                self.data.to_csv(payment_summary_file, index=False)
                self.logger.info(f"Payment summary saved to: {payment_summary_file}")
            
            self.logger.info("Payment data preparation completed")
            
            return True
            
        except Exception as e:
            self.logger.error(f"Payment data preparation error: {str(e)}")
            return False
            
    def generate_reports(self):
        """Generate processing reports."""
        self.update_progress(85, "Generating reports...")
        
        try:
            self.logger.info("Generating processing reports...")
            
            # Create reports directory
            reports_dir = Path("reports")
            reports_dir.mkdir(exist_ok=True)
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # Generate processing summary
            summary = {
                "processing_date": datetime.now().isoformat(),
                "input_file": self.csv_file_path,
                "payments_processed": self.process_payments.get(),
                "total_records": len(self.data) if self.data is not None else 0,
                "status": "completed_successfully"
            }
            
            summary_file = reports_dir / f"processing_summary_{timestamp}.json"
            with open(summary_file, 'w') as f:
                json.dump(summary, f, indent=2)
                
            self.logger.info(f"Processing summary saved to: {summary_file}")
            
            return True
            
        except Exception as e:
            self.logger.error(f"Report generation error: {str(e)}")
            return False
            
    def start_processing(self):
        """Main processing function."""
        try:
            # Validate inputs
            if not self.validate_input():
                return
                
            self.process_button.config(state='disabled')
            self.logger.info("="*50)
            self.logger.info("Starting LDCC1 data processing")
            self.logger.info(f"Processing payments: {self.process_payments.get()}")
            self.logger.info("="*50)
            
            # Load and validate data
            if not self.load_data():
                return
                
            if not self.validate_data_structure():
                return
                
            # Process benefits
            if not self.process_benefits():
                return
                
            # Process reconciliation
            if not self.process_reconciliation():
                return
                
            # Handle payments if selected
            if self.process_payments.get():
                if not self.prepare_payment_data():
                    return
                    
                self.update_progress(90, "Payment processing completed - Ready for EQ online")
                self.logger.info("="*50)
                self.logger.info("PAYMENT PROCESSING COMPLETED")
                self.logger.info("Data is ready for EQ online banking")
                self.logger.info("Please proceed to EQ online to complete payments")
                self.logger.info("="*50)
                
                messagebox.showinfo(
                    "Processing Complete", 
                    "Payment processing completed successfully!\n\n" +
                    "The system has processed all data up to the EQ online stage.\n" +
                    "Please log into EQ online banking to complete the payment process."
                )
            else:
                self.update_progress(90, "Processing completed (no payments)")
                self.logger.info("Processing completed successfully (payments not selected)")
                
            # Generate reports
            if not self.generate_reports():
                return
                
            self.update_progress(100, "All processing completed successfully")
            self.logger.info("LDCC1 data processing completed successfully")
            
            if not self.process_payments.get():
                messagebox.showinfo("Processing Complete", "Data processing completed successfully!")
                
        except Exception as e:
            self.logger.error(f"Processing failed: {str(e)}")
            self.logger.error(f"Traceback:\n{traceback.format_exc()}")
            messagebox.showerror("Processing Error", f"Processing failed:\n{str(e)}")
            
        finally:
            self.process_button.config(state='normal')
            
    def clear_log(self):
        """Clear the log text area."""
        self.log_text.delete(1.0, tk.END)
        
    def save_log(self):
        """Save the current log to a file."""
        try:
            log_content = self.log_text.get(1.0, tk.END)
            
            file_path = filedialog.asksaveasfilename(
                title="Save Log File",
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
            )
            
            if file_path:
                with open(file_path, 'w') as f:
                    f.write(log_content)
                messagebox.showinfo("Success", f"Log saved to: {file_path}")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save log:\n{str(e)}")
            
    def run(self):
        """Run the application."""
        self.logger.info("Starting LDCC1 Processor GUI")
        self.root.mainloop()


def main():
    """Main entry point."""
    try:
        # Check Python version
        if sys.version_info < (3, 6):
            print("Error: This script requires Python 3.6 or higher")
            sys.exit(1)
            
        # Create and run application
        app = LDCC1Processor()
        app.run()
        
    except Exception as e:
        print(f"Fatal error: {e}")
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()