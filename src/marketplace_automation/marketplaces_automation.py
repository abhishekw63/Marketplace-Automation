import pandas as pd
from datetime import datetime
import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from config import Config
from utils import format_indian
from email_service import EmailService

from marketplaces.blinkit import process_blinkit
from marketplaces.flipkart import process_flipkart
from marketplaces.swiggy import process_swiggy

class POReportApp:
    """
    A GUI application to generate Purchase Order (PO) reports for different marketplaces.
    Supports: Blinkit, Flipkart, Swiggy, Zepto (coming soon)
    Includes email functionality to send summary reports.
    """
    
    def __init__(self, root):
        """Initialize the main window and configure the UI."""
        self.root = root
        self.root.title("PO Report Generator")
        self.root.geometry("500x700")
        self.root.resizable(False, False)
        self.root.configure(bg="#667eea") # Simulated gradient background
        self.marketplace_var = tk.StringVar(value="Blinkit")
        self.selected_file_path = None
        self.last_summary = {}
        self.is_processing = False
        
        # Check expiration before building UI
        if not self._check_expiration():
            self.root.destroy()
            return
        
        self._build_ui()
    
    def _check_expiration(self):
        """Check if the application has expired."""
        try:
            expiry_date = datetime.strptime(Config.EXPIRY_DATE, "%d-%m-%Y").date()
            current_date = datetime.now().date()
            
            if current_date > expiry_date:
                messagebox.showerror(
                    "Application Expired",
                    f"This application expired on {Config.EXPIRY_DATE}.\n\n"
                    f"Please contact the administrator for an updated version.\n\n"
                    f"Owner: RENEE-723"
                )
                return False
            
            # Show warning if expiring within 7 days
            days_remaining = (expiry_date - current_date).days
            if days_remaining <= 7:
                messagebox.showwarning(
                    "Expiration Warning",
                    f"⚠️ This application will expire in {days_remaining} day(s).\n\n"
                    f"Expiry Date: {Config.EXPIRY_DATE}\n\n"
                    f"Please contact the administrator for renewal."
                )
            
            return True
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to check expiration: {str(e)}")
            return False
    
    def _build_ui(self):
        """Construct the Tkinter widgets for the application."""
        # Main Container mimicking the HTML container
        container = tk.Frame(self.root, bg="white", highlightthickness=0, bd=0)
        container.pack(fill="both", expand=True, padx=20, pady=20)

        # Rounded corners illusion using relief/border isn't perfect in pure tk,
        # but container provides the white card over gradient background.

        # Header Frame
        header_frame = tk.Frame(container, bg="#667eea", pady=30, bd=0)
        header_frame.pack(fill="x", side="top")

        tk.Label(header_frame, text="📊", font=("Segoe UI", 36), bg="#667eea", fg="white").pack(pady=(0, 10))
        tk.Label(header_frame, text="PO Report Generator", font=("Segoe UI", 20, "bold"), bg="#667eea", fg="white").pack()
        tk.Label(header_frame, text="Automated Purchase Order Intelligence System", font=("Segoe UI", 10), bg="#667eea", fg="#e2e8f0").pack()

        # Content Frame
        content_frame = tk.Frame(container, bg="white", padx=25, pady=25)
        content_frame.pack(fill="both", expand=True)

        # 1. Marketplace Selection
        tk.Label(content_frame, text="SELECT MARKETPLACE", font=("Segoe UI", 10, "bold"), bg="white", fg="#2d3748").pack(anchor="w", pady=(0, 10))

        grid_frame = tk.Frame(content_frame, bg="white")
        grid_frame.pack(fill="x", pady=(0, 20))

        for i in range(4):
            grid_frame.columnconfigure(i, weight=1, pad=10)

        marketplaces = [
            ("Blinkit", "🚀", 0, 0),
            ("Amazon", "🛒", 0, 1),
            ("Flipkart", "📦", 0, 2),
            ("Meesho", "🎯", 0, 3)
        ]

        self.cards = {}
        for name, icon, row, col in marketplaces:
            card = tk.Frame(grid_frame, bg="#f7fafc", bd=1, relief="solid", highlightbackground="#e2e8f0", highlightthickness=1)
            card.grid(row=row, column=col, sticky="nsew", padx=5)

            lbl_icon = tk.Label(card, text=icon, font=("Segoe UI", 20), bg="#f7fafc")
            lbl_icon.pack(pady=(10, 0))
            lbl_name = tk.Label(card, text=name, font=("Segoe UI", 9, "bold"), bg="#f7fafc", fg="#2d3748")
            lbl_name.pack(pady=(5, 10))

            # Bind clicks
            for w in (card, lbl_icon, lbl_name):
                w.bind("<Button-1>", lambda e, m=name: self._select_marketplace(m))

            self.cards[name] = {
                "frame": card,
                "icon": lbl_icon,
                "name": lbl_name
            }

        # Initialize default selection visually
        self._select_marketplace("Blinkit")

        # 2. File Upload Area
        tk.Label(content_frame, text="UPLOAD DATA FILE", font=("Segoe UI", 10, "bold"), bg="white", fg="#2d3748").pack(anchor="w", pady=(0, 10))

        self.upload_area = tk.Frame(content_frame, bg="#f7fafc", bd=1, highlightbackground="#cbd5e0", highlightthickness=2, highlightcolor="#667eea", cursor="hand2")
        self.upload_area.pack(fill="x", pady=(0, 10))

        # Bind dashed border style (using dashes is hard in Frame, so we use solid with color)

        self.upload_icon = tk.Label(self.upload_area, text="📁", font=("Segoe UI", 24), bg="#f7fafc", fg="#667eea")
        self.upload_icon.pack(pady=(15, 5))

        self.upload_text = tk.Label(self.upload_area, text="Click to browse or drag & drop", font=("Segoe UI", 11, "bold"), bg="#f7fafc", fg="#2d3748")
        self.upload_text.pack()

        self.upload_hint = tk.Label(self.upload_area, text="(CSV, XLSX)", font=("Segoe UI", 9), bg="#f7fafc", fg="#718096")
        self.upload_hint.pack(pady=(0, 15))

        for w in (self.upload_area, self.upload_icon, self.upload_text, self.upload_hint):
            w.bind("<Button-1>", self._select_file)

        # Selected file label (hidden initially)
        self.file_info_frame = tk.Frame(content_frame, bg="#f0fff4", bd=1, highlightbackground="#9ae6b4", highlightthickness=1)
        self.file_info_label = tk.Label(self.file_info_frame, text="✓ No file selected", font=("Segoe UI", 10), bg="#f0fff4", fg="#22543d")
        self.file_info_label.pack(side="left", padx=10, pady=8)

        # 3. Generate Button & Status
        self.generate_btn = tk.Button(
            content_frame,
            text="GENERATE REPORT",
            font=("Segoe UI", 11, "bold"),
            bg="#667eea",
            fg="white",
            bd=0,
            activebackground="#764ba2",
            activeforeground="white",
            cursor="hand2",
            command=self.generate_report
        )
        self.generate_btn.pack(fill="x", pady=(20, 10), ipady=8)

        # Progress bar
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TProgressbar", thickness=4, background="#667eea", troughcolor="white")
        self.progress_bar = ttk.Progressbar(content_frame, mode='indeterminate', style="TProgressbar")

        self.status_var = tk.StringVar(value="")
        self.status_label = tk.Label(
            content_frame,
            textvariable=self.status_var,
            font=("Segoe UI", 9),
            bg="white",
            fg="#667eea"
        )
        self.status_label.pack()

        # Coming soon
        tk.Label(
            content_frame,
            text="✨ More marketplaces coming soon!",
            font=("Segoe UI", 9, "bold"),
            bg="#f0fff4",
            fg="#48bb78",
            padx=10,
            pady=8
        ).pack(fill="x", pady=(10, 0))

        # Footer Frame
        footer_frame = tk.Frame(container, bg="#f7fafc", bd=1, highlightbackground="#e2e8f0", highlightthickness=1)
        footer_frame.pack(fill="x", side="bottom")
        
        tk.Label(
            footer_frame,
            text="👨‍💻 Developer: Abhishek Wagh",
            font=("Segoe UI", 9, "bold"),
            bg="#f7fafc",
            fg="#718096"
        ).pack(pady=(10, 2))
        
        tk.Label(
            footer_frame,
            text="📍 Owner ID: RENEE-723\n📧 abhishek.wagh@reneecosmetics.in",
            font=("Segoe UI", 8),
            bg="#f7fafc",
            fg="#667eea"
        ).pack(pady=(0, 10))
    
    def _select_marketplace(self, marketplace):
        """Handle marketplace card selection."""
        self.marketplace_var.set(marketplace)

        # Reset all cards
        for name, widgets in self.cards.items():
            bg_color = "#f7fafc"
            bd_color = "#e2e8f0"
            widgets["frame"].configure(bg=bg_color, highlightbackground=bd_color)
            widgets["icon"].configure(bg=bg_color)
            widgets["name"].configure(bg=bg_color)

        # Highlight selected card
        if marketplace in self.cards:
            selected_widgets = self.cards[marketplace]
            bg_color = "#edf2f7" # Slightly darker background for selected state
            bd_color = "#667eea" # Highlight border color

            selected_widgets["frame"].configure(bg=bg_color, highlightbackground=bd_color)
            selected_widgets["icon"].configure(bg=bg_color)
            selected_widgets["name"].configure(bg=bg_color)

    def _select_file(self, event=None):
        """Handle file selection from the upload area."""
        file_path = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            self.selected_file_path = file_path
            filename = os.path.basename(file_path)
            self.file_info_label.config(text=f"✓ {filename}")

            # Show the selected file info frame (by packing it below upload area)
            self.file_info_frame.pack(fill="x", pady=(0, 10), before=self.generate_btn)

            # Visual feedback on upload area
            self.upload_area.configure(bg="#edf2f7", highlightbackground="#667eea")
            self.upload_icon.configure(bg="#edf2f7")
            self.upload_text.configure(bg="#edf2f7")
            self.upload_hint.configure(bg="#edf2f7")

    def calculate_summary_data(self, df, marketplace):
        """Calculate summary statistics from the DataFrame."""
        if marketplace == "Blinkit":
            total_pos = df['po_number'].nunique()
            total_units = df['units_ordered'].sum()
            total_value = df['total_amount'].sum()
            df['order_date_dt'] = pd.to_datetime(df['order_date'], format='%d-%m-%Y', errors='coerce')
            df['expiry_date_dt'] = pd.to_datetime(df['expiry_date'], format='%d-%m-%Y', errors='coerce')
            min_date = df['order_date_dt'].min()
            max_date = df['expiry_date_dt'].max()
        elif marketplace in ["Flipkart", "Swiggy"]:
            total_pos = df['PO'].nunique()
            total_units = df['PO Qty'].sum()

            # Create a copy to avoid modifying original dataframe and raising SettingWithCopyWarning
            df_calc = df.copy()
            if df_calc['PO Value'].dtype == 'O': # Object/String type
                 df_calc['PO Value_num'] = df_calc['PO Value'].replace('[₹,]', '', regex=True).astype(float)
            else:
                 df_calc['PO Value_num'] = df_calc['PO Value']
            total_value = df_calc['PO Value_num'].sum()

            min_date = df['Order Date'].min()
            max_date = df['Expiry Date'].max()
        else:
            total_pos = df.shape[0]
            total_units = 0
            total_value = 0
            min_date = max_date = None
        
        min_date_str = min_date.strftime('%d-%m-%Y') if hasattr(min_date, 'strftime') else str(min_date)
        max_date_str = max_date.strftime('%d-%m-%Y') if hasattr(max_date, 'strftime') else str(max_date)
        
        return {
            'total_pos': total_pos,
            'total_units': total_units, 
            'total_value': total_value,
            'min_date': min_date_str,
            'max_date': max_date_str
        }

    def generate_report(self):
        """Main function to trigger report generation in a background thread."""
        if self.is_processing:
            return

        marketplace = self.marketplace_var.get()
        if marketplace in ["Amazon", "Meesho", "Zepto"]:
            messagebox.showinfo("Info", f"{marketplace} integration is coming soon!")
            return

        file_path = self.selected_file_path
        if not file_path:
            messagebox.showwarning("No File Selected", "Please select a file to generate the report.")
            return

        self.is_processing = True
        self.generate_btn.config(state=tk.DISABLED, text="GENERATING...")
        self.status_var.set(f"Processing {marketplace} data...")

        self.progress_bar.pack(fill="x", pady=(0, 10), before=self.status_label)
        self.progress_bar.start(10)

        self.root.update_idletasks()

        # Run processing in a background thread
        thread = threading.Thread(target=self._process_report_thread, args=(marketplace, file_path))
        thread.daemon = True
        thread.start()

    def _process_report_thread(self, marketplace, file_path):
        """Thread worker for processing the report."""
        try:
            if marketplace == "Blinkit":
                result = process_blinkit(file_path)
            elif marketplace == "Flipkart":
                result = process_flipkart(file_path)
            elif marketplace == "Swiggy":
                result = process_swiggy(file_path)
            else:
                self.root.after(0, self._handle_process_error, "Unknown marketplace selected.")
                return
                
            # Calculate summary data for UI and Email
            self.last_summary = self.calculate_summary_data(result['tracker_df'], marketplace)
            
            # Update UI on success
            self.root.after(0, self._handle_process_success, result)

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.root.after(0, self._handle_process_error, str(e))

    def _handle_process_error(self, error_msg):
        """Handle errors from the processing thread on the main UI thread."""
        self.is_processing = False
        self.generate_btn.config(state=tk.NORMAL, text="GENERATE REPORT")
        self.progress_bar.stop()
        self.progress_bar.pack_forget()
        self.status_var.set("")
        messagebox.showerror("Error", f"Something went wrong:\n{error_msg}")

    def _handle_process_success(self, result):
        """Handle successful processing and prompt user interactions."""
        try:
            self.is_processing = False
            self.generate_btn.config(state=tk.NORMAL, text="GENERATE REPORT")
            self.progress_bar.stop()
            self.progress_bar.pack_forget()
            self.status_var.set("")

            marketplace = result['marketplace']
            output_file = result['output_file']
            has_sku_data = result['has_sku_data']
            sku_count = result['sku_count']

            # Build success message
            success_msg = f"Report created successfully at:\n{output_file}"
            if marketplace == "Swiggy":
                 output_folder = result.get('output_folder')
                 success_msg = f"Main report created: {output_file.name}\n\nIndividual PO workbooks created in:\n{output_folder.name}"

            # Show summary popup
            sku_status = f"SKU Data: ✅ Available ({sku_count} SKUs)" if has_sku_data else f"SKU Data: ❌ Not Available ({marketplace} format)"

            summary_text = (
                f"📊 SUMMARY REPORT 📊\n\n"
                f"Total POs: {self.last_summary['total_pos']}\n"
                f"Order Date Range: {self.last_summary['min_date']} to {self.last_summary['max_date']}\n"
                f"Total Units Ordered: {format_indian(self.last_summary['total_units'])}\n"
                f"Total Order Value: ₹ {format_indian(self.last_summary['total_value'])}\n"
                f"{sku_status}"
            )

            messagebox.showinfo("Success", success_msg)
            messagebox.showinfo("PO Summary", summary_text)

            # Ask if user wants to send email
            if messagebox.askyesno("Send Email", "Do you want to send this summary via email?"):
                self.status_var.set("Sending email...")
                self.root.update_idletasks()

                # We can run email sending synchronously here or in a thread,
                # for simplicity we'll block the UI briefly since it's an end-step
                success, err_msg = EmailService.send_email_summary(
                    marketplace,
                    self.last_summary,
                    result['tracker_df'],
                    result['sku_df']
                )

                self.status_var.set("")

                if success:
                    recipient_info = f"To: {Config.DEFAULT_RECIPIENT}"
                    if Config.CC_RECIPIENTS:
                        cc_list = "\n".join([f"  • {email}" for email in Config.CC_RECIPIENTS])
                        recipient_info += f"\n\nCC:\n{cc_list}"
                    messagebox.showinfo("Email Sent", f"Summary sent successfully to:\n\n{recipient_info}")
            else:
                messagebox.showerror("Email Error", f"Failed to send email:\n{err_msg}")

            # Ask to open files
            if marketplace == "Swiggy":
                if messagebox.askyesno("Open File", "Do you want to open the main Excel report?"):
                    os.startfile(output_file)
                if messagebox.askyesno("Open Folder", "Do you want to open the PO workbooks folder?"):
                    os.startfile(result['output_folder'])
            else:
                if messagebox.askyesno("Open File", "Do you want to open the generated Excel file?"):
                    os.startfile(output_file)
                    self.root.destroy()

            if marketplace != "Swiggy":
                # original logic destroyed root for blinkit/flipkart if they opened file
                pass
        
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.is_processing = False
            self.generate_btn.config(state=tk.NORMAL, text="GENERATE REPORT")
            self.progress_bar.stop()
            self.progress_bar.pack_forget()
            self.status_var.set("")
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")

def main():
    root = tk.Tk()
    app = POReportApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()