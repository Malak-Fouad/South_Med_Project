import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os


class DarkExcelStopProcessor:
    def __init__(self, root):

        self.root = root
        self.root.title("Excel Stop Processor - Dark Theme")
        self.root.geometry("700x550")
        self.root.configure(bg='#2b2b2b')

        self.file_path = None
        self.data = None
        self.lineroutes_data = None
        self.sheets = []

        self.configure_dark_theme() 
        self.create_widgets()
    
    def configure_dark_theme(self):
        style = ttk.Style()
        bg_color = '#2b2b2b'
        dark_grey = '#3c3f41'
        medium_grey = '#555555'
        light_grey = '#bbbbbb'
        accent_color = '#4ec9b0'
        
        style.theme_use('default')
        
        # Configure styles for different widgets
        style.configure('TFrame', background=bg_color)
        style.configure('TLabel', background=bg_color, foreground=light_grey, font=('Arial', 10))
        style.configure('TButton', background=dark_grey, foreground=light_grey, 
                       font=('Arial', 10, 'bold'), borderwidth=1, focusthickness=3, focuscolor='none')
        style.configure('TLabelframe', background=bg_color, foreground=accent_color, font=('Arial', 11, 'bold'))
        style.configure('TLabelframe.Label', background=bg_color, foreground=accent_color)
        style.configure('TCombobox', fieldbackground=dark_grey, background=dark_grey, 
                       foreground=light_grey, selectbackground=medium_grey)
        
        # Treeview style
        style.configure('Treeview', 
                       background=dark_grey,
                       foreground=light_grey,
                       fieldbackground=dark_grey,
                       font=('Arial', 9))
        
        style.configure('Treeview.Heading', 
                       background=medium_grey,
                       foreground=light_grey,
                       font=('Arial', 10, 'bold'))
        
        style.map('Treeview', background=[('selected', '#0d3d56')])
        style.map('TButton', background=[('active', '#555555')])
        style.map('TCombobox', fieldbackground=[('readonly', dark_grey)], 
                 background=[('readonly', dark_grey)])
    
    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title
        title_label = tk.Label(main_frame, text="South Med", 
                              font=('Arial', 16, 'bold'), 
                              bg='#2b2b2b', fg='#4ec9b0')
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # File selection frame
        file_frame = ttk.LabelFrame(main_frame, text="📁 Select Excel File", padding="12")
        file_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))
        
        self.file_label = ttk.Label(file_frame, text="No file selected", font=('Arial', 9))
        self.file_label.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 15))
        
        browse_btn = ttk.Button(file_frame, text="Browse Excel File", 
                               command=self.browse_file)
        browse_btn.grid(row=0, column=1)
        
        # Sheet info frame
        self.sheet_info_frame = ttk.LabelFrame(main_frame, text="📄 Detected Sheets", padding="12")
        self.sheet_info_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))
        self.sheet_info_frame.grid_remove()  # Hide initially
        
        # Line Route Item sheet info
        lri_label = ttk.Label(self.sheet_info_frame, text="Lineroute items sheet:")
        lri_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.lri_sheet_label = ttk.Label(self.sheet_info_frame, text="Not detected", foreground="#4ec9b0")
        self.lri_sheet_label.grid(row=0, column=1, sticky=tk.W)
        
        # Lineroutes sheet info
        lr_label = ttk.Label(self.sheet_info_frame, text="Lineroutes sheet:")
        lr_label.grid(row=1, column=0, sticky=tk.W, padx=(0, 10))
        self.lr_sheet_label = ttk.Label(self.sheet_info_frame, text="Not detected", foreground="#4ec9b0")
        self.lr_sheet_label.grid(row=1, column=1, sticky=tk.W)
        
        # Process button
        self.process_btn = ttk.Button(main_frame, text="🚀 PROCESS DATA", 
                                     command=self.process_data, state="disabled")
        self.process_btn.grid(row=3, column=0, columnspan=2, pady=15, ipadx=20, ipady=5)
        
        # Results frame
        results_frame = ttk.LabelFrame(main_frame, text="📊 PROCESSED RESULTS", padding="12")
        results_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        
        # Treeview for results
        self.tree = ttk.Treeview(results_frame, 
                                columns=("$LINEROUTEITEM:LINENAME", "LINEROUTENAME", "Stops", "HubName", "LINKRUNTIME", "MAX:LINEROUTEITEMS\\VOL(AP)"), 
                                show="headings", height=12)
        self.tree.heading("$LINEROUTEITEM:LINENAME", text="LINE NAME")
        self.tree.heading("LINEROUTENAME", text="LINE ROUTE NAME")
        self.tree.heading("Stops", text="STOPS ARRAY")
        self.tree.heading("HubName", text="HUB NAME")
        self.tree.heading("LINKRUNTIME", text="LINK RUNTIME")
        self.tree.heading("MAX:LINEROUTEITEMS\\VOL(AP)", text="MAX VOL(AP)")
        
        self.tree.column("$LINEROUTEITEM:LINENAME", width=120)
        self.tree.column("LINEROUTENAME", width=150)
        self.tree.column("Stops", width=150)
        self.tree.column("HubName", width=80)
        self.tree.column("LINKRUNTIME", width=100)
        self.tree.column("MAX:LINEROUTEITEMS\\VOL(AP)", width=100)
        
        # Scrollbar for treeview
        scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Export button
        self.export_btn = ttk.Button(main_frame, text="💾 EXPORT RESULTS", 
                                    command=self.export_results, state="disabled")
        self.export_btn.grid(row=5, column=0, columnspan=2, pady=10, ipadx=15, ipady=4)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready to process Excel file")
        status_bar = tk.Label(main_frame, textvariable=self.status_var, 
                             font=('Arial', 8), bg='#2b2b2b', fg='#888888')
        status_bar.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(4, weight=1)
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        file_frame.columnconfigure(0, weight=1)
        self.sheet_info_frame.columnconfigure(1, weight=1)
    
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if file_path:
            self.file_path = file_path
            filename = os.path.basename(file_path)
            self.file_label.config(text=filename)
            
            # Clear previous results
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            # Get available sheets
            try:
                self.sheets = pd.ExcelFile(file_path).sheet_names
                self.status_var.set(f"Found {len(self.sheets)} sheets in file")
                
                # Detect which sheets to use
                line_route_item_sheet = None
                lineroutes_sheet = None
                
                # Look for Line Route Item sheet
                for sheet_name in self.sheets:
                    if "lineroute items" in sheet_name.lower():
                        line_route_item_sheet = sheet_name
                        break
                
                # If not found, try alternative names
                if not line_route_item_sheet:
                    for sheet_name in self.sheets:
                        if "lineroute item" in sheet_name.lower():
                            line_route_item_sheet = sheet_name
                            break
                
                # Look for Lineroutes sheet
                for sheet_name in self.sheets:
                    if sheet_name.lower() == "lineroutes":
                        lineroutes_sheet = sheet_name
                        break
                
                # If not found, try alternative names
                if not lineroutes_sheet:
                    for sheet_name in self.sheets:
                        if "lineroutes" in sheet_name.lower() and sheet_name != line_route_item_sheet:
                            lineroutes_sheet = sheet_name
                            break
                
                # Update sheet info labels
                self.lri_sheet_label.config(text=line_route_item_sheet or "Not found")
                self.lr_sheet_label.config(text=lineroutes_sheet or "Not found")
                
                # Show sheet info frame
                self.sheet_info_frame.grid()
                
                if line_route_item_sheet and lineroutes_sheet:
                    self.status_var.set(f"Ready to process: {line_route_item_sheet} + {lineroutes_sheet}")
                    self.process_btn.config(state="normal")
                else:
                    self.status_var.set("Could not detect required sheets")
                    self.process_btn.config(state="disabled")
                    
            except Exception as e:
                messagebox.showerror("Error", f"Could not read Excel file: {str(e)}")
                self.status_var.set("Error reading Excel file")
                self.process_btn.config(state="disabled")
    
    def extract_hub_name(self, stop_names):
        # Filter out empty/NaN values and take the first valid one
        valid_stop_names = [name for name in stop_names if pd.notna(name) and str(name).strip() != '']
        
        if valid_stop_names:
            first_stop_name = str(valid_stop_names[0])
            
            if 'Ext. Hub01' in first_stop_name:
                return 'Ext. Hub01'
            elif 'Ext. Hub02' in first_stop_name:
                return 'Ext. Hub02'
            elif 'Gate3' in first_stop_name:
                return 'Gate3'
        
        return 'Unknown Hub'
    
    def format_stop_numbers(self, stop_numbers):
        """Convert stop numbers to integers and remove .0 decimal points"""
        formatted_stops = []
        for stop in stop_numbers:
            try:
                # Convert to integer to remove decimal points
                formatted_stop = str(int(float(stop))) if pd.notna(stop) else ''
                formatted_stops.append(formatted_stop)
            except (ValueError, TypeError):
                # If conversion fails, use original value
                formatted_stops.append(str(stop))
        return formatted_stops
    
    def process_data(self):
        if not self.file_path:
            messagebox.showerror("Error", "Please select an Excel file first!")
            return
        
        try:
            self.status_var.set("Processing data...")
            self.root.update()
            
            # Detect sheets again to ensure we have the right ones
            line_route_item_sheet = None
            lineroutes_sheet = None
            
            # Look for Line Route Item sheet
            for sheet_name in self.sheets:
                if "lineroute items" in sheet_name.lower():
                    line_route_item_sheet = sheet_name
                    break
            
            # If not found, try alternative names
            if not line_route_item_sheet:
                for sheet_name in self.sheets:
                    if "lineroute item" in sheet_name.lower():
                        line_route_item_sheet = sheet_name
                        break
            
            # Look for Lineroutes sheet
            for sheet_name in self.sheets:
                if sheet_name.lower() == "lineroutes":
                    lineroutes_sheet = sheet_name
                    break
            
            # If not found, try alternative names
            if not lineroutes_sheet:
                for sheet_name in self.sheets:
                    if "lineroutes" in sheet_name.lower() and sheet_name != line_route_item_sheet:
                        lineroutes_sheet = sheet_name
                        break
            
            if not line_route_item_sheet or not lineroutes_sheet:
                messagebox.showerror("Error", "Could not detect required sheets in the Excel file!")
                return
            
            # Read both sheets
            self.data = pd.read_excel(self.file_path, sheet_name=line_route_item_sheet)
            self.lineroutes_data = pd.read_excel(self.file_path, sheet_name=lineroutes_sheet)
            
            # Check for required columns in Line Route Item sheet - handle different column names
            required_columns_variations = [
                ['$LINEROUTEITEM:LINENAME', 'LINEROUTENAME', 'STOPPOINTNO'],  # original
                ['$LINEROUTEITEM:LINENAME', 'LINEROUTENAME', 'SSTOPPOINT:NO']  # new variation
            ]

            stop_name_columns = ['STOPPOINT\\NAME', 'STOPPOINT/NAME', 'STOPPOINT NAME',]

            # Find the correct set of columns
            missing_columns = []
            stop_point_col = None

            for column_set in required_columns_variations:
                missing_columns = [col for col in column_set if col not in self.data.columns]
                if not missing_columns:
                    # Found matching column set
                    if 'STOPPOINTNO' in column_set:
                        stop_point_col = 'STOPPOINTNO'
                    else:
                        stop_point_col = 'SSTOPPOINT:NO'
                    break

            if missing_columns:
                messagebox.showerror("Error", f"Missing columns in Line Route Item sheet: {', '.join(missing_columns)}")
                self.status_var.set("Error: Missing columns in Excel file")
                return

            # Find the correct stop name column
            stop_name_col = None
            for col in stop_name_columns:
                if col in self.data.columns:
                    stop_name_col = col
                    break

            if not stop_name_col:
                messagebox.showerror("Error", f"Could not find stop name column. Available columns: {list(self.data.columns)}")
                self.status_var.set("Error: Missing stop name column")
                return
            
            # Check for required columns in Lineroutes sheet
            lineroutes_required_columns = ['NAME', 'LINKRUNTIME', 'MAX:LINEROUTEITEMS\\VOL(AP)']
            lineroutes_missing_columns = [col for col in lineroutes_required_columns if col not in self.lineroutes_data.columns]
            
            if lineroutes_missing_columns:
                messagebox.showerror("Error", f"Missing columns in Lineroutes sheet: {', '.join(lineroutes_missing_columns)}")
                self.status_var.set("Error: Missing columns in Lineroutes sheet")
                return

            for item in self.tree.get_children():
                self.tree.delete(item)
            
            self.status_var.set("Removing null values and duplicates...")
            self.root.update()
            
            initial_count = len(self.data)
            data_clean = self.data.dropna(subset=[stop_point_col])
            null_removed = initial_count - len(data_clean)
            
            # Remove duplicates
            data_clean = data_clean.drop_duplicates(subset=['$LINEROUTEITEM:LINENAME', 'LINEROUTENAME', stop_point_col])
            duplicates_removed = (initial_count - null_removed) - len(data_clean)
            
            if null_removed > 0:
                self.status_var.set(f"Removed {null_removed} null entries and {duplicates_removed} duplicates")
                self.root.update()
            elif duplicates_removed > 0:
                self.status_var.set(f"Removed {duplicates_removed} duplicate entries")
                self.root.update()
            
            # First, get hub name for each LineName (same for all routes in the same line)
            line_hubs = {}
            for line_name in data_clean['$LINEROUTEITEM:LINENAME'].unique():
                line_data = data_clean[data_clean['$LINEROUTEITEM:LINENAME'] == line_name]
                # Get all stop names for this line and extract hub from first valid one
                all_stop_names = line_data[stop_name_col].dropna().tolist()
                hub_name = self.extract_hub_name(all_stop_names)
                line_hubs[line_name] = hub_name
            
            grouped_data = data_clean.groupby(['$LINEROUTEITEM:LINENAME', 'LINEROUTENAME']).agg({
                stop_point_col: list,
                stop_name_col: list
            }).reset_index()
            
            # Add HubName column using the line-level hub mapping
            grouped_data['HubName'] = grouped_data['$LINEROUTEITEM:LINENAME'].map(line_hubs)
            
            # Merge with Lineroutes data based on LINEROUTENAME = NAME
            merged_data = grouped_data.merge(
                self.lineroutes_data[['NAME', 'LINKRUNTIME', 'MAX:LINEROUTEITEMS\\VOL(AP)']],
                left_on='LINEROUTENAME',
                right_on='NAME',
                how='left'
            )
            
            # Add data to treeview
            for _, row in merged_data.iterrows():
                # Format stop numbers to remove .0
                formatted_stops = self.format_stop_numbers(row[stop_point_col])
                stops_array = ' → '.join(formatted_stops)
                self.tree.insert("", "end", values=(
                    row['$LINEROUTEITEM:LINENAME'], 
                    row['LINEROUTENAME'], 
                    stops_array, 
                    row['HubName'],
                    row['LINKRUNTIME'],
                    row['MAX:LINEROUTEITEMS\\VOL(AP)']
                ))
            
            total_removed = null_removed + duplicates_removed
            self.status_var.set(f"Successfully processed {len(merged_data)} unique LineRouteNames (removed {null_removed} null + {duplicates_removed} duplicates)")
            self.export_btn.config(state="normal")  # Enable export button
            messagebox.showinfo("Success", f"Processed {len(merged_data)} unique LineRouteNames!\nRemoved {null_removed} null entries and {duplicates_removed} duplicate entries.")
            
        except Exception as e:
            error_msg = f"Error processing file: {str(e)}"
            self.status_var.set(f"Error: {str(e)}")
            messagebox.showerror("Error", error_msg)
    
    def export_results(self):
        if self.data is None or self.data.empty:
            messagebox.showerror("Error", "No data to export! Please process files first.")
            return
        
        try:
            self.status_var.set("Exporting results...")
            self.root.update()
            
            # Check for required columns in Line Route Item sheet - handle different column names
            required_columns_variations = [
                ['$LINEROUTEITEM:LINENAME', 'LINEROUTENAME', 'STOPPOINTNO'],  # original
                ['$LINEROUTEITEM:LINENAME', 'LINEROUTENAME', 'SSTOPPOINT:NO']  # new variation
            ]

            stop_name_columns = ['STOPPOINT\\NAME', 'STOPPOINT/NAME', 'STOPPOINT NAME']

            # Find the correct set of columns
            missing_columns = []
            stop_point_col = None

            for column_set in required_columns_variations:
                missing_columns = [col for col in column_set if col not in self.data.columns]
                if not missing_columns:
                    # Found matching column set
                    if 'STOPPOINTNO' in column_set:
                        stop_point_col = 'STOPPOINTNO'
                    else:
                        stop_point_col = 'SSTOPPOINT:NO'
                    break

            if missing_columns:
                messagebox.showerror("Error", f"Missing columns in Line Route Item sheet: {', '.join(missing_columns)}")
                return

            # Find the correct stop name column
            stop_name_col = None
            for col in stop_name_columns:
                if col in self.data.columns:
                    stop_name_col = col
                    break

            if not stop_name_col:
                messagebox.showerror("Error", "Could not find stop name column in data")
                return
            
            # Remove rows with null values in StopPointNo
            data_clean = self.data.dropna(subset=[stop_point_col])
            # Remove duplicates
            data_clean = data_clean.drop_duplicates(subset=['$LINEROUTEITEM:LINENAME', 'LINEROUTENAME', stop_point_col])
            
            # First, get hub name for each LineName (same for all routes in the same line)
            line_hubs = {}
            for line_name in data_clean['$LINEROUTEITEM:LINENAME'].unique():
                line_data = data_clean[data_clean['$LINEROUTEITEM:LINENAME'] == line_name]
                # Get all stop names for this line and extract hub from first valid one
                all_stop_names = line_data[stop_name_col].dropna().tolist()
                hub_name = self.extract_hub_name(all_stop_names)
                line_hubs[line_name] = hub_name
            
            # Group and aggregate data with hub name
            grouped_data = data_clean.groupby(['$LINEROUTEITEM:LINENAME', 'LINEROUTENAME']).agg({
                stop_point_col: list,
                stop_name_col: list
            }).reset_index()
            
            grouped_data['HubName'] = grouped_data['$LINEROUTEITEM:LINENAME'].map(line_hubs)
            
            # Merge with Lineroutes data
            merged_data = grouped_data.merge(
                self.lineroutes_data[['NAME', 'LINKRUNTIME', 'MAX:LINEROUTEITEMS\\VOL(AP)']],
                left_on='LINEROUTENAME',
                right_on='NAME',
                how='left'
            )
            
            output_data = []
            for _, row in merged_data.iterrows():
                # Format stop numbers to remove .0
                formatted_stops = self.format_stop_numbers(row[stop_point_col])
                stops_string = ' → '.join(formatted_stops)
                output_data.append({
                    '$LINEROUTEITEM:LINENAME': row['$LINEROUTEITEM:LINENAME'],
                    'LINEROUTENAME': row['LINEROUTENAME'],
                    'StopsArray': stops_string,
                    'HubName': row['HubName'],
                    'LINKRUNTIME': row['LINKRUNTIME'],
                    'MAX:LINEROUTEITEMS\\VOL(AP)': row['MAX:LINEROUTEITEMS\\VOL(AP)']
                })
            
            output_df = pd.DataFrame(output_data)
            
            # Save to Excel
            output_file = filedialog.asksaveasfilename(
                title="Save Results As",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
            
            if output_file:
                output_df.to_excel(output_file, index=False)
                self.status_var.set(f"Results exported to {os.path.basename(output_file)}")
                messagebox.showinfo("Success", f"Results exported to {output_file}")
                
        except Exception as e:
            error_msg = f"Error exporting results: {str(e)}"
            self.status_var.set(f"Export error: {str(e)}")
            messagebox.showerror("Error", error_msg)

def main():
    root = tk.Tk()
    
    try:
        root.iconbitmap(default='')  
    except:
        pass
    
    app = DarkExcelStopProcessor(root)
    root.mainloop()

if __name__ == "__main__":
    main()