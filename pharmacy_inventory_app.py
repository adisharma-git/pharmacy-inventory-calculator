"""
Pharmacy Inventory Management System
Desktop Application for Windows

This application automates pharmacy inventory calculations by:
- Loading master data, consumption data, and pending orders
- Calculating stock levels, reorder needs, and optimal order quantities
- Generating an Excel output file with all calculations

Author: Claude
Version: 1.0
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
from pathlib import Path
import traceback
from datetime import datetime


class PharmacyInventoryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Pharmacy Inventory Calculator")
        self.root.geometry("700x550")
        self.root.resizable(False, False)
        
        # File paths
        self.master_file = tk.StringVar()
        self.consumption_file = tk.StringVar()
        self.expected_file = tk.StringVar()
        self.output_file = tk.StringVar(value="Inventory_Calculation.xlsx")
        
        # Set default output location to Desktop
        desktop = Path.home() / "Desktop"
        self.output_file.set(str(desktop / "Inventory_Calculation.xlsx"))
        
        self.setup_ui()
    
    def setup_ui(self):
        """Create the user interface"""
        # Title
        title_frame = tk.Frame(self.root, bg="#2C3E50", height=80)
        title_frame.pack(fill=tk.X)
        title_frame.pack_propagate(False)
        
        title_label = tk.Label(
            title_frame, 
            text="Pharmacy Inventory Calculator",
            font=("Arial", 20, "bold"),
            bg="#2C3E50",
            fg="white"
        )
        title_label.pack(pady=25)
        
        # Main content frame
        content_frame = tk.Frame(self.root, padx=30, pady=20)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # Instructions
        instructions = tk.Label(
            content_frame,
            text="Select the three input Excel files and click Process to generate the inventory report.",
            font=("Arial", 10),
            wraplength=600,
            justify=tk.LEFT
        )
        instructions.pack(pady=(0, 20))
        
        # File selection section
        self.create_file_selector(
            content_frame, 
            "Master Data File:", 
            self.master_file,
            row=0
        )
        
        self.create_file_selector(
            content_frame, 
            "Drug Consumption File:", 
            self.consumption_file,
            row=1
        )
        
        self.create_file_selector(
            content_frame, 
            "Expected Items File:", 
            self.expected_file,
            row=2
        )
        
        # Separator
        separator = ttk.Separator(content_frame, orient='horizontal')
        separator.pack(fill=tk.X, pady=20)
        
        # Output file section
        output_label = tk.Label(
            content_frame,
            text="Output File:",
            font=("Arial", 10, "bold")
        )
        output_label.pack(anchor=tk.W, pady=(0, 5))
        
        output_frame = tk.Frame(content_frame)
        output_frame.pack(fill=tk.X, pady=(0, 20))
        
        output_entry = tk.Entry(
            output_frame,
            textvariable=self.output_file,
            font=("Arial", 9),
            state="readonly"
        )
        output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        output_btn = tk.Button(
            output_frame,
            text="Change Location",
            command=self.select_output_location,
            bg="#95A5A6",
            fg="white",
            font=("Arial", 9),
            relief=tk.FLAT,
            padx=15,
            pady=5
        )
        output_btn.pack(side=tk.LEFT)
        
        # Process button
        self.process_btn = tk.Button(
            content_frame,
            text="PROCESS INVENTORY",
            command=self.process_inventory,
            bg="#27AE60",
            fg="white",
            font=("Arial", 12, "bold"),
            relief=tk.FLAT,
            padx=30,
            pady=15,
            cursor="hand2"
        )
        self.process_btn.pack(pady=10)
        
        # Status label
        self.status_label = tk.Label(
            content_frame,
            text="Ready to process",
            font=("Arial", 9),
            fg="#7F8C8D"
        )
        self.status_label.pack(pady=10)
        
        # Progress bar
        self.progress = ttk.Progressbar(
            content_frame,
            mode='indeterminate',
            length=500
        )
        
    def create_file_selector(self, parent, label_text, var, row):
        """Create a file selector row"""
        frame = tk.Frame(parent)
        frame.pack(fill=tk.X, pady=5)
        
        label = tk.Label(
            frame,
            text=label_text,
            font=("Arial", 10, "bold"),
            width=20,
            anchor=tk.W
        )
        label.pack(side=tk.LEFT)
        
        entry = tk.Entry(
            frame,
            textvariable=var,
            font=("Arial", 9),
            state="readonly"
        )
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        btn = tk.Button(
            frame,
            text="Browse",
            command=lambda: self.browse_file(var),
            bg="#3498DB",
            fg="white",
            font=("Arial", 9),
            relief=tk.FLAT,
            padx=15,
            pady=5
        )
        btn.pack(side=tk.LEFT)
    
    def browse_file(self, var):
        """Open file browser dialog"""
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            var.set(filename)
    
    def select_output_location(self):
        """Select output file location"""
        filename = filedialog.asksaveasfilename(
            title="Save Output File As",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile="Inventory_Calculation.xlsx"
        )
        if filename:
            self.output_file.set(filename)
    
    def update_status(self, message, color="#7F8C8D"):
        """Update status message"""
        self.status_label.config(text=message, fg=color)
        self.root.update()
    
    def validate_files(self):
        """Validate that all required files are selected"""
        if not self.master_file.get():
            messagebox.showerror("Missing File", "Please select the Master Data file.")
            return False
        if not self.consumption_file.get():
            messagebox.showerror("Missing File", "Please select the Drug Consumption file.")
            return False
        if not self.expected_file.get():
            messagebox.showerror("Missing File", "Please select the Expected Items file.")
            return False
        
        # Check if files exist
        for file_path in [self.master_file.get(), self.consumption_file.get(), self.expected_file.get()]:
            if not Path(file_path).exists():
                messagebox.showerror("File Not Found", f"File not found:\n{file_path}")
                return False
        
        return True
    
    def process_inventory(self):
        """Main processing function"""
        if not self.validate_files():
            return
        
        try:
            # Disable button and show progress
            self.process_btn.config(state=tk.DISABLED)
            self.progress.pack(pady=10)
            self.progress.start(10)
            
            self.update_status("Loading files...", "#3498DB")
            
            # Load data
            master_df = pd.read_excel(self.master_file.get())
            consumption_df = pd.read_excel(self.consumption_file.get())
            expected_df = pd.read_excel(self.expected_file.get())
            
            self.update_status("Processing inventory calculations...", "#3498DB")
            
            # Process the data
            result_df = self.calculate_inventory(master_df, consumption_df, expected_df)
            
            self.update_status("Saving output file...", "#3498DB")
            
            # Save output
            result_df.to_excel(self.output_file.get(), index=False)
            
            # Success
            self.progress.stop()
            self.progress.pack_forget()
            self.process_btn.config(state=tk.NORMAL)
            self.update_status("✓ Processing completed successfully!", "#27AE60")
            
            # Show success message
            result = messagebox.askyesno(
                "Success",
                f"Inventory calculation completed!\n\n"
                f"Output saved to:\n{self.output_file.get()}\n\n"
                f"Would you like to open the file?",
                icon='info'
            )
            
            if result:
                import os, sys, subprocess
                if sys.platform == "win32":
                    os.startfile(self.output_file.get())
                elif sys.platform == "darwin":
                    subprocess.call(["open", self.output_file.get()])
                else:
                    subprocess.call(["xdg-open", self.output_file.get()])
            
        except Exception as e:
            self.progress.stop()
            self.progress.pack_forget()
            self.process_btn.config(state=tk.NORMAL)
            self.update_status("✗ Error occurred during processing", "#E74C3C")
            
            error_msg = f"An error occurred:\n\n{str(e)}\n\n{traceback.format_exc()}"
            messagebox.showerror("Processing Error", error_msg)
    
    def calculate_inventory(self, master_df, consumption_df, expected_df):
        """
        Perform all inventory calculations
        
        Returns: DataFrame with all calculated columns
        """
        # Start with master data (columns A-P)
        # Copy only the first 16 columns (A-P as specified)
        result_df = master_df.iloc[:, :16].copy()
        
        # Normalize column names for easier handling
        consumption_df.columns = consumption_df.columns.str.strip()
        expected_df.columns = expected_df.columns.str.strip()
        
        # Create lookup dictionaries for better performance
        # Global Stock lookup (Column O in consumption file)
        global_stock_dict = dict(zip(
            consumption_df['Drug Code'].fillna(''),
            consumption_df['Global Stock'].fillna(0)
        ))
        
        # Main Store Stock lookup (Column M in consumption file - "Local\nStock")
        main_store_col = [col for col in consumption_df.columns if 'Local' in col and 'Stock' in col][0]
        main_store_dict = dict(zip(
            consumption_df['Drug Code'].fillna(''),
            consumption_df[main_store_col].fillna(0)
        ))
        
        # Pending PO lookup (Column M in expected file - "Pend.\nQty")
        pending_col = [col for col in expected_df.columns if 'Pend' in col][0]
        pending_dict = dict(zip(
            expected_df['Drug \nCode'].fillna(''),
            expected_df[pending_col].fillna(0)
        ))
        
        # Column Q - Global Stock
        result_df['Global Stock'] = result_df['Drug Code'].map(global_stock_dict).fillna(0)
        
        # Column R - Main Store Stock
        result_df['Main Store Stock'] = result_df['Drug Code'].map(main_store_dict).fillna(0)
        
        # Column S - Pending PO
        result_df['Pending PO'] = result_df['Drug Code'].map(pending_dict).fillna(0).astype(int)
        
        # Column T - Net Stock
        result_df['Net Stock'] = result_df['Global Stock'] + result_df['Pending PO']
        
        # Column U - Global Stock Days
        result_df['Global Stock Days'] = np.where(
            result_df['ADC'] > 0,
            result_df['Global Stock'] / result_df['ADC'],
            0
        )
        
        # Column V - Main Store Stock Days
        result_df['Main Store Stock Days'] = np.where(
            result_df['ADC'] > 0,
            result_df['Main Store Stock'] / result_df['ADC'],
            0
        )
        
        # Column W - Reorder Needed
        # Only reorder if Current SKU is True AND Net Stock < Min Stock Level
        result_df['Reorder Needed'] = np.where(
            (result_df['Current SKU '] == True) & 
            (result_df['Net Stock'] < pd.to_numeric(result_df['Min Stock Level'], errors='coerce').fillna(0)),
            1.0,
            0.0
        )
        
        # Column X - Order Qty
        # Calculate: ROUNDUP((Max - Net) / Pack Size, 0) * Pack Size
        max_stock = pd.to_numeric(result_df['Max Stock Level'], errors='coerce').fillna(0)
        
        result_df['Order Qty'] = np.where(
            result_df['Reorder Needed'] == 1.0,
            np.ceil((max_stock - result_df['Net Stock']) / result_df['Pack Size']) * result_df['Pack Size'],
            0.0
        )
        
        # Ensure Order Qty is not negative
        result_df['Order Qty'] = result_df['Order Qty'].clip(lower=0)
        
        return result_df


def main():
    """Application entry point"""
    root = tk.Tk()
    app = PharmacyInventoryApp(root)
    
    # Center window on screen
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    root.mainloop()


if __name__ == "__main__":
    main()