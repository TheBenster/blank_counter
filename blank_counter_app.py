import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import threading
from collections import defaultdict
import os
import sys

class BlankCounterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("📊 Excel Blank Counter - Modern Edition")
        self.root.geometry("950x750")
        
        # Modern light theme colors
        self.colors = {
            'bg': '#f8fafc',           # Very light gray background
            'surface': '#ffffff',       # White cards
            'surface_variant': '#f1f5f9', # Light gray variant
            'primary': '#3b82f6',      # Modern blue
            'primary_dark': '#2563eb', # Darker blue for hover
            'secondary': '#64748b',    # Slate gray
            'text_primary': '#0f172a', # Dark slate
            'text_secondary': '#64748b', # Medium slate
            'success': '#10b981',      # Green
            'warning': '#f59e0b',      # Amber
            'error': '#ef4444',        # Red
            'border': '#e2e8f0',      # Light border
            'shadow': '#00000008'      # Very light shadow
        }
        
        self.root.configure(bg=self.colors['bg'])
        
        # Configure modern ttk style
        self.setup_modern_styles()
        
        # Data storage
        self.input_data = None
        self.headers = None
        self.transformed_data = None
        self.blank_summary = None
        
        self.setup_ui()
    
    def setup_modern_styles(self):
        """Configure modern ttk styles"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure button style
        style.configure('Modern.TButton',
                       background=self.colors['primary'],
                       foreground='white',
                       borderwidth=0,
                       focuscolor='none',
                       padding=(20, 12))
        
        style.map('Modern.TButton',
                  background=[('active', self.colors['primary_dark']),
                             ('pressed', self.colors['primary_dark'])])
        
        # Configure combobox style
        style.configure('Modern.TCombobox',
                       fieldbackground=self.colors['surface'],
                       background=self.colors['surface'],
                       borderwidth=1,
                       relief='solid',
                       padding=10)
        
        # Configure progressbar style
        style.configure('Modern.Horizontal.TProgressbar',
                       background=self.colors['primary'],
                       troughcolor=self.colors['surface_variant'],
                       borderwidth=0,
                       lightcolor=self.colors['primary'],
                       darkcolor=self.colors['primary'])
    
    def setup_ui(self):
        # Main scrollable container
        canvas = tk.Canvas(self.root, bg=self.colors['bg'], highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=self.colors['bg'])
        
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        
        # Pack canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Bind mousewheel to canvas
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)  # Windows/Mac
        canvas.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))  # Linux
        canvas.bind_all("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))   # Linux
        
        # Main container with padding
        main_frame = tk.Frame(scrollable_frame, bg=self.colors['bg'])
        main_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=30)
        
        # Header section
        self.create_header(main_frame)
        
        # Cards container
        cards_container = tk.Frame(main_frame, bg=self.colors['bg'])
        cards_container.pack(fill=tk.BOTH, expand=True)
        
        # Create cards
        self.create_info_card(cards_container)
        self.create_file_upload_card(cards_container)
        self.create_column_selection_card(cards_container)
        self.create_preview_card(cards_container)
        self.create_download_card(cards_container)
        
        # Update canvas scroll region
        scrollable_frame.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))
    
    def create_header(self, parent):
        """Create modern header section"""
        header_frame = tk.Frame(parent, bg=self.colors['bg'])
        header_frame.pack(fill=tk.X, pady=(0, 40))
        
        # Title and subtitle
        title_label = tk.Label(header_frame, 
                              text="📊 Excel Blank Counter", 
                              font=("SF Pro Display", 28, "bold"), 
                              bg=self.colors['bg'], 
                              fg=self.colors['text_primary'])
        title_label.pack(anchor=tk.W)
        
        subtitle_label = tk.Label(header_frame,
                                 text="Transform your Excel data with intelligent blank counting",
                                 font=("SF Pro Text", 14),
                                 bg=self.colors['bg'],
                                 fg=self.colors['text_secondary'])
        subtitle_label.pack(anchor=tk.W, pady=(5, 0))
    
    def create_modern_card(self, parent, title, subtitle=None):
        """Create a modern card-style frame"""
        # Card container with shadow effect
        card_container = tk.Frame(parent, bg=self.colors['bg'])
        card_container.pack(fill=tk.X, pady=(0, 24))
        
        # Main card frame
        card_frame = tk.Frame(card_container, 
                             bg=self.colors['surface'],
                             relief="flat",
                             bd=0)
        card_frame.pack(fill=tk.X, padx=2, pady=2)
        
        # Add visual depth with border
        card_frame.configure(highlightbackground=self.colors['border'],
                           highlightthickness=1)
        
        # Content frame with padding
        content_frame = tk.Frame(card_frame, bg=self.colors['surface'])
        content_frame.pack(fill=tk.BOTH, expand=True, padx=32, pady=28)
        
        # Header section
        header_frame = tk.Frame(content_frame, bg=self.colors['surface'])
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Title
        title_label = tk.Label(header_frame, 
                              text=title, 
                              font=("SF Pro Display", 18, "bold"), 
                              bg=self.colors['surface'], 
                              fg=self.colors['text_primary'])
        title_label.pack(anchor=tk.W)
        
        # Subtitle if provided
        if subtitle:
            subtitle_label = tk.Label(header_frame,
                                     text=subtitle,
                                     font=("SF Pro Text", 13),
                                     bg=self.colors['surface'],
                                     fg=self.colors['text_secondary'])
            subtitle_label.pack(anchor=tk.W, pady=(4, 0))
        
        return content_frame
    
    def create_info_card(self, parent):
        content = self.create_modern_card(parent, 
                                         "How This Tool Works",
                                         "Understanding the blank counting logic")
        
        # Info points with modern styling
        info_points = [
            ("✅ Counts as blank:", "Empty cells, null values, and empty strings"),
            ("❌ Does NOT count:", "Zeros, error messages, or any text/numbers"),
            ("🔍 Process:", "Analyzes each person's rows for blank fields"),
            ("📊 Output:", "Long-form table showing blank counts per person per field")
        ]
        
        for emoji_label, description in info_points:
            point_frame = tk.Frame(content, bg=self.colors['surface'])
            point_frame.pack(fill=tk.X, pady=6)
            
            emoji_lbl = tk.Label(point_frame,
                                text=emoji_label,
                                font=("SF Pro Text", 13, "bold"),
                                bg=self.colors['surface'],
                                fg=self.colors['primary'])
            emoji_lbl.pack(side=tk.LEFT)
            
            desc_lbl = tk.Label(point_frame,
                               text=description,
                               font=("SF Pro Text", 13),
                               bg=self.colors['surface'],
                               fg=self.colors['text_secondary'])
            desc_lbl.pack(side=tk.LEFT, padx=(8, 0))
    
    def create_file_upload_card(self, parent):
        content = self.create_modern_card(parent,
                                         "Upload Your File",
                                         "Select the Excel file containing your case data")
        
        # File selection frame
        file_frame = tk.Frame(content, bg=self.colors['surface'])
        file_frame.pack(fill=tk.X, pady=10)
        
        # File path display
        path_frame = tk.Frame(file_frame, 
                             bg=self.colors['surface_variant'],
                             relief="flat",
                             bd=1)
        path_frame.pack(fill=tk.X, pady=(0, 12))
        
        self.file_path_var = tk.StringVar(value="No file selected")
        file_label = tk.Label(path_frame, 
                             textvariable=self.file_path_var,
                             font=("SF Pro Text", 12),
                             bg=self.colors['surface_variant'],
                             fg=self.colors['text_secondary'],
                             anchor="w")
        file_label.pack(fill=tk.X, padx=16, pady=12)
        
        # Browse button
        browse_btn = tk.Button(file_frame, 
                              text="Choose File",
                              command=self.browse_file,
                              bg=self.colors['primary'],
                              fg="white",
                              font=("SF Pro Text", 12, "bold"),
                              relief="flat",
                              bd=0,
                              padx=24,
                              pady=12,
                              cursor="hand2")
        browse_btn.pack(anchor=tk.W)
        
        # Hover effects
        def on_enter(e):
            browse_btn.config(bg=self.colors['primary_dark'])
        def on_leave(e):
            browse_btn.config(bg=self.colors['primary'])
        
        browse_btn.bind("<Enter>", on_enter)
        browse_btn.bind("<Leave>", on_leave)
        
        # Status label
        self.file_status_var = tk.StringVar()
        self.file_status_label = tk.Label(content, 
                                         textvariable=self.file_status_var,
                                         font=("SF Pro Text", 12),
                                         bg=self.colors['surface'])
        self.file_status_label.pack(anchor=tk.W, pady=(10, 0))
    
    def create_column_selection_card(self, parent):
        content = self.create_modern_card(parent,
                                         "Select Assignment Column",
                                         "Choose the column containing person assignments")
        
        # Dropdown container
        dropdown_frame = tk.Frame(content, bg=self.colors['surface'])
        dropdown_frame.pack(fill=tk.X, pady=10)
        
        self.column_var = tk.StringVar()
        self.column_combo = ttk.Combobox(dropdown_frame,
                                        textvariable=self.column_var,
                                        state="disabled",
                                        font=("SF Pro Text", 12),
                                        style="Modern.TCombobox")
        self.column_combo.pack(fill=tk.X, ipady=8)
        self.column_combo.bind("<<ComboboxSelected>>", self.on_column_selected)
    
    def create_preview_card(self, parent):
        content = self.create_modern_card(parent,
                                         "Generate Preview",
                                         "Process your data and see the results")
        
        # Preview button
        self.preview_btn = tk.Button(content,
                                    text="Generate Preview",
                                    command=self.preview_data,
                                    state="disabled",
                                    bg=self.colors['secondary'],
                                    fg="white",
                                    font=("SF Pro Text", 12, "bold"),
                                    relief="flat",
                                    bd=0,
                                    padx=24,
                                    pady=12,
                                    cursor="hand2")
        self.preview_btn.pack(fill=tk.X, pady=(0, 20))
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(content,
                                           variable=self.progress_var,
                                           maximum=100,
                                           style="Modern.Horizontal.TProgressbar",
                                           length=400)
        self.progress_bar.pack(fill=tk.X, pady=(0, 20))
        self.progress_bar.pack_forget()  # Hide initially
        
        # Preview text area
        preview_container = tk.Frame(content, bg=self.colors['surface'])
        preview_container.pack(fill=tk.BOTH, expand=True)
        
        # Text area with modern styling
        text_frame = tk.Frame(preview_container, 
                             bg=self.colors['surface_variant'],
                             relief="flat",
                             bd=1)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        self.preview_text = tk.Text(text_frame,
                                   height=16,
                                   font=("SF Mono", 11),
                                   wrap=tk.WORD,
                                   state="disabled",
                                   bg=self.colors['surface_variant'],
                                   fg=self.colors['text_primary'],
                                   relief="flat",
                                   bd=0,
                                   padx=16,
                                   pady=16,
                                   selectbackground=self.colors['primary'],
                                   selectforeground="white")
        
        scrollbar_preview = ttk.Scrollbar(text_frame, 
                                         orient="vertical",
                                         command=self.preview_text.yview)
        self.preview_text.configure(yscrollcommand=scrollbar_preview.set)
        
        self.preview_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_preview.pack(side=tk.RIGHT, fill=tk.Y)
    
    def create_download_card(self, parent):
        content = self.create_modern_card(parent,
                                         "Download Results",
                                         "Save your transformed data as an Excel file")
        
        self.download_btn = tk.Button(content,
                                     text="💾 Download Excel File",
                                     command=self.download_results,
                                     state="disabled",
                                     bg=self.colors['success'],
                                     fg="white",
                                     font=("SF Pro Text", 12, "bold"),
                                     relief="flat",
                                     bd=0,
                                     padx=24,
                                     pady=12,
                                     cursor="hand2")
        self.download_btn.pack(fill=tk.X)
        
        # Hover effects for download button
        def on_enter_download(e):
            if self.download_btn['state'] != 'disabled':
                self.download_btn.config(bg='#059669')  # Darker green
        def on_leave_download(e):
            if self.download_btn['state'] != 'disabled':
                self.download_btn.config(bg=self.colors['success'])
        
        self.download_btn.bind("<Enter>", on_enter_download)
        self.download_btn.bind("<Leave>", on_leave_download)
    
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if file_path:
            self.file_path_var.set(os.path.basename(file_path))
            self.load_file(file_path)
    
    def load_file(self, file_path):
        try:
            # Read Excel file
            df = pd.read_excel(file_path)
            self.input_data = df
            self.headers = list(df.columns)
            
            # Populate column dropdown
            self.column_combo.config(state="normal")
            self.column_combo['values'] = self.headers
            
            # Try to auto-select "Assigned" column
            for i, header in enumerate(self.headers):
                if 'assigned' in header.lower():
                    self.column_combo.current(i)
                    self.column_var.set(header)
                    self.update_preview_button_state()
                    break
            
            self.column_combo.config(state="readonly")
            
            # Update status
            self.file_status_var.set(f"✅ Successfully loaded {len(df):,} rows and {len(self.headers)} columns")
            self.file_status_label.config(fg=self.colors['success'])
            
        except Exception as e:
            self.file_status_var.set(f"❌ Error: {str(e)}")
            self.file_status_label.config(fg=self.colors['error'])
            messagebox.showerror("Error", f"Error reading file: {str(e)}")
    
    def on_column_selected(self, event=None):
        self.update_preview_button_state()
    
    def update_preview_button_state(self):
        if self.column_var.get():
            self.preview_btn.config(state="normal", bg=self.colors['primary'])
            # Add hover effects for enabled button
            def on_enter_preview(e):
                self.preview_btn.config(bg=self.colors['primary_dark'])
            def on_leave_preview(e):
                self.preview_btn.config(bg=self.colors['primary'])
            
            self.preview_btn.bind("<Enter>", on_enter_preview)
            self.preview_btn.bind("<Leave>", on_leave_preview)
    
    def preview_data(self):
        if not self.column_var.get():
            messagebox.showwarning("Warning", "Please select an 'Assigned To' column first.")
            return
        
        # Show progress bar and update button
        self.progress_bar.pack(fill=tk.X, pady=(0, 20))
        self.progress_var.set(0)
        self.preview_btn.config(state="disabled", text="Processing...", bg=self.colors['secondary'])
        
        # Run transformation in separate thread
        thread = threading.Thread(target=self.transform_data_thread)
        thread.daemon = True
        thread.start()
    
    def transform_data_thread(self):
        try:
            assigned_to_col = self.column_var.get()
            
            # Update progress
            self.root.after(0, lambda: self.progress_var.set(10))
            
            # Transform data
            self.transformed_data, self.blank_summary = self.transform_data(assigned_to_col)
            
            # Update progress
            self.root.after(0, lambda: self.progress_var.set(90))
            
            # Display preview
            self.root.after(0, self.display_preview)
            
            # Complete
            self.root.after(0, lambda: self.progress_var.set(100))
            self.root.after(500, self.hide_progress)
            
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", f"Error processing data: {str(e)}"))
            self.root.after(0, self.hide_progress)
    
    def transform_data(self, assigned_to_col):
        df = self.input_data.copy()
        results = []
        total_blanks = 0
        blanks_by_person = defaultdict(int)
        blanks_by_field = defaultdict(int)
        
        # Handle unassigned (blank) values
        df[assigned_to_col] = df[assigned_to_col].fillna('(Unassigned)')
        df.loc[df[assigned_to_col] == '', assigned_to_col] = '(Unassigned)'
        
        unique_assigned_to = df[assigned_to_col].unique()
        
        for person in unique_assigned_to:
            person_rows = df[df[assigned_to_col] == person]
            
            for col in df.columns:
                blank_count = 0
                
                for _, row in person_rows.iterrows():
                    cell_value = row[col]
                    # Check if truly blank (matching Excel's (Blanks) filter)
                    is_blank = (
                        pd.isna(cell_value) or
                        cell_value == '' or
                        (isinstance(cell_value, str) and cell_value.strip() == '')
                    )
                    
                    if is_blank:
                        blank_count += 1
                        total_blanks += 1
                
                results.append([person, col, blank_count])
                blanks_by_person[person] += blank_count
                blanks_by_field[col] += blank_count
        
        summary = {
            'total_blanks': total_blanks,
            'blanks_by_person': dict(blanks_by_person),
            'blanks_by_field': dict(blanks_by_field),
            'total_persons': len(unique_assigned_to),
            'total_fields': len(df.columns)
        }
        
        return results, summary
    
    def display_preview(self):
        if not self.transformed_data or not self.blank_summary:
            return
        
        self.preview_text.config(state="normal")
        self.preview_text.delete(1.0, tk.END)
        
        # Modern summary display
        summary = self.blank_summary
        summary_text = f"""📊 ANALYSIS RESULTS
{'─' * 60}

📈 Overview
Total Blanks Found: {summary['total_blanks']:,}
People Analyzed: {summary['total_persons']}
Fields Analyzed: {summary['total_fields']}
Output Rows Generated: {len(self.transformed_data):,}

👥 Top Contributors (Most Blanks)
{'─' * 40}
"""
        
        # Top people with modern formatting
        top_people = sorted(summary['blanks_by_person'].items(), key=lambda x: x[1], reverse=True)[:5]
        for i, (person, count) in enumerate(top_people, 1):
            summary_text += f"{i:2}. {person:<25} {count:>6} blanks\n"
        
        summary_text += f"\n📋 Most Problematic Fields\n{'─' * 40}\n"
        
        # Top fields with modern formatting
        top_fields = sorted(summary['blanks_by_field'].items(), key=lambda x: x[1], reverse=True)[:5]
        for i, (field, count) in enumerate(top_fields, 1):
            field_short = field[:22] + "..." if len(field) > 25 else field
            summary_text += f"{i:2}. {field_short:<25} {count:>6} blanks\n"
        
        summary_text += f"\n📋 Data Preview (First 50 Rows)\n{'─' * 60}\n"
        summary_text += f"{'Assigned To':<22} {'Field':<25} {'Count':>8}\n"
        summary_text += f"{'─' * 22} {'─' * 25} {'─' * 8}\n"
        
        # Preview data with better formatting
        for row in self.transformed_data[:50]:
            assigned_to = str(row[0])[:20] + "..." if len(str(row[0])) > 20 else str(row[0])
            field = str(row[1])[:23] + "..." if len(str(row[1])) > 25 else str(row[1])
            summary_text += f"{assigned_to:<22} {field:<25} {row[2]:>8}\n"
        
        if len(self.transformed_data) > 50:
            summary_text += f"\n... and {len(self.transformed_data) - 50:,} more rows\n"
        
        summary_text += f"\n✅ Ready for download!"
        
        self.preview_text.insert(1.0, summary_text)
        self.preview_text.config(state="disabled")
        
        # Enable download button
        self.download_btn.config(state="normal")
    
    def hide_progress(self):
        self.progress_bar.pack_forget()
        self.preview_btn.config(state="normal", text="Generate Preview", bg=self.colors['primary'])
    
    def download_results(self):
        if not self.transformed_data:
            messagebox.showwarning("Warning", "Please generate the preview first.")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="Save Transformed Data",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialname="blank_count_analysis.xlsx"
        )
        
        if file_path:
            try:
                # Create DataFrame with headers
                df_output = pd.DataFrame(self.transformed_data,
                                       columns=['Assigned To', 'Field', 'Blank_Count'])
                
                # Save to Excel
                df_output.to_excel(file_path, index=False)
                
                messagebox.showinfo("Success", 
                                   f"✅ File saved successfully!\n\nLocation: {file_path}\nRows: {len(df_output):,}")
                
            except Exception as e:
                messagebox.showerror("Error", f"❌ Error saving file:\n{str(e)}")

def main():
    root = tk.Tk()
    
    # Try to set app icon (works on Windows/Linux)
    try:
        # You can replace this with your icon file path
        # root.iconbitmap('icon.ico')  # Uncomment and add your icon path
        pass
    except:
        pass
    
    app = BlankCounterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()