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
        self.root.geometry("460x320")
        self.root.resizable(False, False)
        self.marketplace_var = tk.StringVar(value="Blinkit")
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
        ttk.Label(self.root, text="PO Report Generator", font=("Segoe UI", 16, "bold")).pack(pady=(20, 10))
        ttk.Label(self.root, text="Select Marketplace:").pack(pady=(10, 5))
        self.marketplace_dropdown = ttk.Combobox(
            self.root,
            textvariable=self.marketplace_var,
            values=["Blinkit", "Flipkart", "Swiggy", "Zepto"],
            state="readonly",
            width=25
        )
        self.marketplace_dropdown.pack()

        self.generate_btn = ttk.Button(
            self.root,
            text="Select CSV/Xlsx and Generate Report",
            command=self.generate_report
        )
        self.generate_btn.pack(pady=(20, 5))

        self.status_var = tk.StringVar(value="")
        self.status_label = ttk.Label(self.root, textvariable=self.status_var, font=("Segoe UI", 9), foreground="blue")
        self.status_label.pack(pady=(0, 10))

        ttk.Label(
            self.root,
            text="Other marketplaces are coming soon!",
            font=("Segoe UI", 9, "italic"),
            foreground="green"
        ).pack()
        ttk.Label(
            self.root,
            text=f"Last Updated: {datetime.now().strftime('%d-%m-%Y')} | Owner: RENEE-723",
            font=("Segoe UI", 8),
            foreground="gray"
        ).pack(pady=(5, 0))
    
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
        if marketplace == "Zepto":
            messagebox.showinfo("Info", "Zepto integration is coming soon!")
            return

        file_path = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx *.xls")]
        )
        if not file_path:
            return

        self.is_processing = True
        self.generate_btn.config(state=tk.DISABLED)
        self.status_var.set(f"Processing {marketplace} data...")
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
            self.last_summary = self.calculate_summary_data(result['df'], marketplace)
            
            # Update UI on success
            self.root.after(0, self._handle_process_success, result)

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.root.after(0, self._handle_process_error, str(e))

    def _handle_process_error(self, error_msg):
        """Handle errors from the processing thread on the main UI thread."""
        self.is_processing = False
        self.generate_btn.config(state=tk.NORMAL)
        self.status_var.set("")
        messagebox.showerror("Error", f"Something went wrong:\n{error_msg}")

    def _handle_process_success(self, result):
        """Handle successful processing and prompt user interactions."""
        self.is_processing = False
        self.generate_btn.config(state=tk.NORMAL)
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

def main():
    root = tk.Tk()
    app = POReportApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()