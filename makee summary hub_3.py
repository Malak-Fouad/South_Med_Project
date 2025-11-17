import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
from pathlib import Path
import threading
from collections import defaultdict

class ExcelHubProcessor:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Hub Processor - Dark Theme")
        self.root.geometry("900x700")
        self.root.configure(bg='#2b2b2b')
        
        # Apply dark theme
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.configure_dark_theme()
        
        self.input_folder = ""
        self.output_folder = ""
        
        self.create_widgets()
        
    def configure_dark_theme(self):
        # Configure dark theme colors
        bg_color = '#2b2b2b'
        fg_color = '#ffffff'
        accent_color = '#404040'
        button_color = '#404040'
        
        self.style.configure('TFrame', background=bg_color)
        self.style.configure('TLabel', background=bg_color, foreground=fg_color, font=('Arial', 10))
        self.style.configure('TButton', background=button_color, foreground=fg_color, 
                           font=('Arial', 10), borderwidth=1)
        self.style.configure('TLabelframe', background=bg_color, foreground=fg_color)
        self.style.configure('TLabelframe.Label', background=bg_color, foreground=fg_color)
        self.style.configure('TProgressbar', background='#007acc', troughcolor=accent_color)
        
        self.style.map('TButton',
                      background=[('active', '#505050'), ('pressed', '#606060')])
        
    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text="Excel Hub Processor - Combine Same HubName", 
                               font=('Arial', 16, 'bold'))
        title_label.pack(pady=(0, 20))
        
        # Description
        desc_label = ttk.Label(main_frame, 
                              text="This tool combines Excel files with the same HubName and sums Hub Area columns",
                              font=('Arial', 10))
        desc_label.pack(pady=(0, 20))
        
        # Input folder selection
        input_frame = ttk.LabelFrame(main_frame, text="Input Folder", padding="10")
        input_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.input_var = tk.StringVar()
        input_entry = ttk.Entry(input_frame, textvariable=self.input_var, width=60)
        input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        input_btn = ttk.Button(input_frame, text="Browse", command=self.select_input_folder)
        input_btn.pack(side=tk.RIGHT)
        
        # Output folder selection
        output_frame = ttk.LabelFrame(main_frame, text="Output Folder", padding="10")
        output_frame.pack(fill=tk.X, pady=(0, 20))
        
        self.output_var = tk.StringVar()
        output_entry = ttk.Entry(output_frame, textvariable=self.output_var, width=60)
        output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        output_btn = ttk.Button(output_frame, text="Browse", command=self.select_output_folder)
        output_btn.pack(side=tk.RIGHT)
        
        # Process button
        self.process_btn = ttk.Button(main_frame, text="Combine Excel Files by HubName", 
                                     command=self.start_processing)
        self.process_btn.pack(pady=(0, 20))
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.pack(fill=tk.X, pady=(0, 20))
        
        # Log text area
        log_frame = ttk.LabelFrame(main_frame, text="Processing Log", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create scrollable text widget
        text_frame = ttk.Frame(log_frame)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = tk.Text(text_frame, height=15, bg='#1e1e1e', fg='#ffffff',
                               insertbackground='white', font=('Consolas', 9))
        
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready to process files")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, 
                              relief=tk.SUNKEN, style='TLabel')
        status_bar.pack(fill=tk.X, pady=(10, 0))
        
    def log_message(self, message):
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def select_input_folder(self):
        folder = filedialog.askdirectory(title="Select Input Folder containing Excel files")
        if folder:
            self.input_folder = folder
            self.input_var.set(folder)
            self.log_message(f"Input folder set: {folder}")
            
    def select_output_folder(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_folder = folder
            self.output_var.set(folder)
            self.log_message(f"Output folder set: {folder}")
            
    def start_processing(self):
        if not self.input_folder or not os.path.exists(self.input_folder):
            messagebox.showerror("Error", "Please select a valid input folder")
            return
            
        if not self.output_folder:
            messagebox.showerror("Error", "Please select an output folder")
            return
            
        # Disable process button and start progress bar
        self.process_btn.config(state='disabled')
        self.progress.start()
        self.status_var.set("Processing files...")
        
        # Run processing in separate thread to keep GUI responsive
        thread = threading.Thread(target=self.process_files)
        thread.daemon = True
        thread.start()
        
    def process_files(self):
        try:
            self.log_message("Starting file processing...")
            
            # Get all Excel files in input folder
            excel_files = []
            for ext in ['*.xlsx', '*.xls']:
                excel_files.extend(Path(self.input_folder).glob(ext))
                
            if not excel_files:
                self.log_message("No Excel files found in the input folder")
                return
                
            self.log_message(f"Found {len(excel_files)} Excel files")
            
            # Dictionary to store files by HubName
            hub_files = defaultdict(list)
            
            # First pass: Identify HubNames in each file
            for file_path in excel_files:
                try:
                    self.log_message(f"Scanning: {file_path.name}")
                    
                    # Read Excel file
                    df = pd.read_excel(file_path)
                    
                    # Check if HubName column exists
                    if 'HubName' not in df.columns:
                        self.log_message(f"Warning: No HubName column in {file_path.name}")
                        continue
                    
                    # Get unique HubNames in this file
                    hub_names = df['HubName'].dropna().unique()
                    
                    for hub_name in hub_names:
                        hub_name = str(hub_name).strip()
                        hub_files[hub_name].append(file_path)
                        
                except Exception as e:
                    self.log_message(f"Error scanning {file_path.name}: {str(e)}")
                    continue
            
            # Process hubs that have multiple files
            processed_hubs = 0
            
            for hub_name, files in hub_files.items():
                if len(files) < 2:
                    continue  # Skip hubs with only one file
                    
                try:
                    self.log_message(f"Processing Hub '{hub_name}' with {len(files)} files")
                    self.process_hub_files(hub_name, files)
                    processed_hubs += 1
                    
                except Exception as e:
                    self.log_message(f"Error processing hub {hub_name}: {str(e)}")
                    continue
            
            if processed_hubs == 0:
                self.log_message("No hubs with multiple files found for processing")
            else:
                self.log_message(f"Successfully processed {processed_hubs} hubs with multiple files")
            
            self.log_message("Processing completed successfully!")
            self.status_var.set("Processing completed")
            
        except Exception as e:
            self.log_message(f"Unexpected error: {str(e)}")
            self.status_var.set("Error occurred")
        finally:
            # Re-enable process button and stop progress bar
            self.root.after(0, self.processing_finished)

    def get_short_name(self, filename):
        """Extract short name from filename - take last 2 words"""
        words = filename.split('_')
        if len(words) >= 1:
            return ' '.join(words[-1:]).title()
        else:
            return filename
    
    def process_hub_files(self, hub_name, files):
        """Process multiple files for the same HubName and combine them"""
        
        # Read all files for this hub
        file_data = []
        for file_path in files:
            try:
                df = pd.read_excel(file_path)
                # Filter rows for this specific hub
                hub_df = df[df['HubName'] == hub_name].copy()
                hub_df['Source_File'] = file_path.stem  # Store filename without extension
                file_data.append((file_path.stem, hub_df))
                self.log_message(f"  - Loaded {len(hub_df)} rows from {file_path.name}")
            except Exception as e:
                self.log_message(f"  - Error reading {file_path.name}: {str(e)}")
                continue
        
        if len(file_data) < 2:
            self.log_message(f"  - Not enough valid files for hub {hub_name}")
            return
        
        # Create base structure with Bus_Capacity and Headway
        base_columns = ['Bus_Capacity', 'Headway (min)']
        combined_df = None
        
        # Process each file and add its specific columns
        for filename, df in file_data:
            # Only process Hub_Area columns (remove Fleet_Size)
            file_specific_cols = ['Hub_Area']  # Only keep Hub_Area, remove Fleet_Size
            
            # Check which columns exist in this file
            available_cols = [col for col in file_specific_cols if col in df.columns]
            
            if not available_cols:
                self.log_message(f"  - No Hub_Area column found in {filename}")
                continue
            
            # Create a temporary dataframe for this file's data
            temp_df = df[base_columns + available_cols].copy()
            
            # Rename columns to include shortened filename
            col_rename = {}
            for col in available_cols:
                short_name = self.get_short_name(filename)
                new_name = f"Hub_Area_{short_name.replace(' ', '_')}"
                col_rename[col] = new_name
            temp_df = temp_df.rename(columns=col_rename)
            
            # Merge with combined dataframe
            if combined_df is None:
                combined_df = temp_df
            else:
                combined_df = pd.merge(combined_df, temp_df, on=base_columns, how='outer')
        
        if combined_df is None or combined_df.empty:
            self.log_message(f"  - No data to combine for hub {hub_name}")
            return
        
        # Calculate sums for Hub Area columns
        hub_area_cols = [col for col in combined_df.columns if 'Hub_Area' in col and 'Sum_' not in col]
        
        if hub_area_cols:
            combined_df['Sum_Hub_Area'] = combined_df[hub_area_cols].sum(axis=1)
        
        # Fill NaN values with 0
        combined_df = combined_df.fillna(0)
        
        # Save the combined file
        output_filename = f"Summary_{hub_name.replace(' ', '_').replace('/', '_')}.xlsx"
        output_path = os.path.join(self.output_folder, output_filename)
        
        combined_df.to_excel(output_path, index=False)
        
        self.log_message(f"  - Created combined file: {output_filename}")
        self.log_message(f"  - Total rows: {len(combined_df)}")
        self.log_message(f"  - Columns: {list(combined_df.columns)}")
        
        # Log column sums
        if 'Sum_Hub_Area_1_Hour' in combined_df.columns:
            self.log_message(f"  - Total Sum_Hub_Area_1_Hour: {combined_df['Sum_Hub_Area_1_Hour'].sum():.2f}")
            
    def processing_finished(self):
        """Called when processing is finished to update UI"""
        self.progress.stop()
        self.process_btn.config(state='normal')
        messagebox.showinfo("Processing Complete", "Excel files have been processed successfully!")

def main():
    root = tk.Tk()
    app = ExcelHubProcessor(root)
    root.mainloop()

if __name__ == "__main__":
    main()