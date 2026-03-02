import pandas as pd
from datetime import datetime
import os
import sys

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QComboBox, QPushButton, QFileDialog, QFrame,
    QSizePolicy, QGraphicsDropShadowEffect
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont, QIcon, QCursor, QColor

from config import Config
from utils import format_indian
from email_service import EmailService
from ui_smooth import SmoothDialog

from marketplaces.blinkit import process_blinkit
from marketplaces.flipkart import process_flipkart
from marketplaces.swiggy import process_swiggy

class ReportWorker(QThread):
    """Background worker thread to handle report processing without freezing UI."""
    success = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, marketplace, file_path):
        super().__init__()
        self.marketplace = marketplace
        self.file_path = file_path

    def run(self):
        try:
            if self.marketplace == "Blinkit":
                result = process_blinkit(self.file_path)
            elif self.marketplace == "Flipkart":
                result = process_flipkart(self.file_path)
            elif self.marketplace == "Swiggy":
                result = process_swiggy(self.file_path)
            else:
                self.error.emit("Unknown marketplace selected.")
                return

            self.success.emit(result)
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.error.emit(str(e))


class POReportApp(QMainWindow):
    """
    A PyQt6 GUI application to generate Purchase Order (PO) reports for different marketplaces.
    Supports: Blinkit, Flipkart, Swiggy, Zepto (coming soon)
    Includes email functionality to send summary reports.
    """
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PO Report Generator")
        self.setFixedSize(600, 500)
        self.last_summary = {}
        self.worker = None

        if not self._check_expiration():
            sys.exit(0)

        # Overall background (iOS-like subtle gray background for main app area)
        self.setStyleSheet("QMainWindow { background-color: #F9FAFB; }")

        self._build_ui()
        self._apply_custom_styles()

    def _check_expiration(self):
        """Check if the application has expired."""
        try:
            expiry_date = datetime.strptime(Config.EXPIRY_DATE, "%d-%m-%Y").date()
            current_date = datetime.now().date()
            
            if current_date > expiry_date:
                SmoothDialog.show_error(
                    self,
                    "Application Expired",
                    f"This application expired on {Config.EXPIRY_DATE}.\n\n"
                    f"Please contact the administrator for an updated version.\n\n"
                    f"Owner: RENEE-723"
                )
                return False
            
            days_remaining = (expiry_date - current_date).days
            if days_remaining <= 7:
                SmoothDialog.show_warning(
                    self,
                    "Expiration Warning",
                    f"⚠️ This application will expire in {days_remaining} day(s).\n\n"
                    f"Expiry Date: {Config.EXPIRY_DATE}\n\n"
                    f"Please contact the administrator for renewal."
                )
            
            return True
        except Exception as e:
            SmoothDialog.show_error(self, "Error", f"Failed to check expiration: {str(e)}")
            return False

    def _build_ui(self):
        """Construct the PyQt6 widgets for the application."""
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # Header Frame
        header_frame = QFrame()
        header_frame.setObjectName("headerFrame")
        header_layout = QVBoxLayout(header_frame)
        header_layout.setContentsMargins(20, 25, 20, 20)
        header_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        title_label = QLabel("📊 PO Report Generator")
        title_label.setObjectName("titleLabel")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        subtitle_label = QLabel("Automated Purchase Order Intelligence System")
        subtitle_label.setObjectName("subtitleLabel")
        subtitle_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        header_layout.addWidget(title_label)
        header_layout.addWidget(subtitle_label)
        main_layout.addWidget(header_frame)

        # Body/Content Frame (Card style)
        content_wrapper = QWidget()
        wrapper_layout = QVBoxLayout(content_wrapper)
        wrapper_layout.setContentsMargins(40, 10, 40, 20)

        card_frame = QFrame()
        card_frame.setObjectName("cardFrame")
        card_layout = QVBoxLayout(card_frame)
        card_layout.setContentsMargins(35, 35, 35, 35)
        card_layout.setSpacing(24)

        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(40)
        shadow.setColor(QColor(0, 0, 0, 20)) # Light smooth shadow
        shadow.setOffset(0, 10)
        card_frame.setGraphicsEffect(shadow)

        # Marketplace Selection
        combo_layout = QVBoxLayout()
        combo_layout.setSpacing(8)

        select_lbl = QLabel("Select Marketplace:")
        select_lbl.setObjectName("selectLabel")

        self.marketplace_dropdown = QComboBox()
        self.marketplace_dropdown.addItems(["Blinkit", "Flipkart", "Swiggy", "Zepto"])
        self.marketplace_dropdown.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))

        combo_layout.addWidget(select_lbl)
        combo_layout.addWidget(self.marketplace_dropdown)

        card_layout.addLayout(combo_layout)

        # Generate Button
        self.generate_btn = QPushButton("📁 Select File & Generate Report")
        self.generate_btn.setObjectName("generateBtn")
        self.generate_btn.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.generate_btn.setMinimumHeight(50)
        self.generate_btn.clicked.connect(self.generate_report)
        card_layout.addWidget(self.generate_btn)

        # Status Label
        self.status_label = QLabel("")
        self.status_label.setObjectName("statusLabel")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        card_layout.addWidget(self.status_label)

        # Info Label
        coming_soon_lbl = QLabel("✨ Other marketplaces are coming soon!")
        coming_soon_lbl.setObjectName("comingSoonLabel")
        coming_soon_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        card_layout.addWidget(coming_soon_lbl)

        wrapper_layout.addWidget(card_frame)
        main_layout.addWidget(content_wrapper)
        
        main_layout.addStretch()

        # Footer Frame
        footer_frame = QFrame()
        footer_frame.setObjectName("footerFrame")
        footer_layout = QVBoxLayout(footer_frame)
        footer_layout.setContentsMargins(10, 15, 10, 15)
        footer_layout.setSpacing(5)

        dev_label = QLabel("👨‍💻 Developer: Abhishek Wagh")
        dev_label.setObjectName("devLabel")
        dev_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        info_label = QLabel("🆔 Owner ID: RENEE-723  •  📧 abhishek.wagh@reneecosmetics.in")
        info_label.setObjectName("infoLabel")
        info_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        footer_layout.addWidget(dev_label)
        footer_layout.addWidget(info_label)
        main_layout.addWidget(footer_frame)
        
    def _apply_custom_styles(self):
        """Apply sleek iOS-like smooth CSS styling."""
        custom_css = """
        * {
            font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif;
        }
        #titleLabel {
            font-size: 28px;
            font-weight: 800;
            color: #111827; /* Dark slate */
        }
        #subtitleLabel {
            font-size: 14px;
            color: #6B7280; /* Neutral gray */
            margin-top: 5px;
        }
        #cardFrame {
            background-color: #FFFFFF;
            border-radius: 24px;
            /* Border removed to let the shadow do the work, like an iOS card */
        }
        #selectLabel {
            font-size: 15px;
            font-weight: 600;
            color: #374151;
        }
        QComboBox {
            padding: 12px 16px;
            font-size: 15px;
            font-weight: 500;
            color: #1F2937;
            background-color: #F3F4F6;
            border-radius: 12px;
            border: 1px solid #E5E7EB;
        }
        QComboBox:hover {
            background-color: #E5E7EB;
            border: 1px solid #D1D5DB;
        }
        QComboBox::drop-down {
            border: none;
            width: 30px;
        }
        QComboBox::down-arrow {
            /* Fallback basic arrow since we don't have custom svgs */
            image: none;
        }
        QComboBox QAbstractItemView {
            background-color: #FFFFFF;
            border-radius: 8px;
            border: 1px solid #E5E7EB;
            selection-background-color: #F3F4F6;
            selection-color: #111827;
            padding: 4px;
            outline: none;
        }
        #generateBtn {
            font-size: 16px;
            font-weight: bold;
            background-color: #007AFF; /* iOS Blue */
            color: white;
            border-radius: 14px;
            border: none;
            padding: 14px;
        }
        #generateBtn:hover {
            background-color: #0066D6;
        }
        #generateBtn:pressed {
            background-color: #0052AC;
        }
        #generateBtn:disabled {
            background-color: #93C5FD;
            color: #EFF6FF;
        }
        #statusLabel {
            font-size: 14px;
            font-weight: 600;
            color: #10B981; /* Emerald green */
        }
        #comingSoonLabel {
            font-size: 13px;
            font-style: italic;
            color: #9CA3AF;
        }
        #footerFrame {
            background-color: #FFFFFF;
            border-top: 1px solid #E5E7EB;
        }
        #devLabel {
            font-size: 14px;
            font-weight: bold;
            color: #4B5563;
        }
        #infoLabel {
            font-size: 12px;
            color: #9CA3AF;
        }
        """
        current_ss = self.styleSheet()
        self.setStyleSheet(current_ss + custom_css)

    def calculate_summary_data(self, df, marketplace):
        """Calculate summary statistics from the DataFrame."""
        if marketplace == "Blinkit":
            total_pos = df['po_number'].nunique()
            total_units = df['units_ordered'].sum()
            total_value = df['total_amount'].sum()
            order_date_dt = pd.to_datetime(df['order_date'], format='%d-%m-%Y', errors='coerce')
            expiry_date_dt = pd.to_datetime(df['expiry_date'], format='%d-%m-%Y', errors='coerce')
            min_date = order_date_dt.min()
            max_date = expiry_date_dt.max()
        elif marketplace in ["Flipkart", "Swiggy"]:
            total_pos = df['PO'].nunique()
            total_units = df['PO Qty'].sum()

            df_calc = df.copy()
            if df_calc['PO Value'].dtype == 'O':
                 df_calc['PO Value_num'] = df_calc['PO Value'].replace('[₹,]', '', regex=True).astype(float)
            else:
                 df_calc['PO Value_num'] = df_calc['PO Value']
            total_value = df_calc['PO Value_num'].sum()

            order_date_dt = pd.to_datetime(df['Order Date'], format='%d-%m-%Y', errors='coerce')
            expiry_date_dt = pd.to_datetime(df['Expiry Date'], format='%d-%m-%Y', errors='coerce')
            min_date = order_date_dt.min()
            max_date = expiry_date_dt.max()
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
        """Main function to trigger report generation via QThread."""
        marketplace = self.marketplace_dropdown.currentText()
        if marketplace == "Zepto":
            SmoothDialog.show_info(self, "Info", "Zepto integration is coming soon!")
            return

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Data File",
            "",
            "Data Files (*.csv *.xlsx *.xls);;CSV files (*.csv);;Excel files (*.xlsx *.xls)"
        )

        if not file_path:
            return

        self.generate_btn.setEnabled(False)
        self.status_label.setText(f"Processing {marketplace} data...")

        # Setup and start background thread
        self.worker = ReportWorker(marketplace, file_path)
        self.worker.success.connect(self._handle_process_success)
        self.worker.error.connect(self._handle_process_error)
        self.worker.start()

    def _handle_process_error(self, error_msg):
        """Handle errors from the processing thread."""
        self.generate_btn.setEnabled(True)
        self.status_label.setText("")
        SmoothDialog.show_error(self, "Error", f"Something went wrong:\n{error_msg}")

    def _handle_process_success(self, result):
        """Handle successful processing and prompt user interactions."""
        self.generate_btn.setEnabled(True)
        self.status_label.setText("")

        marketplace = result['marketplace']
        output_file = result['output_file']
        has_sku_data = result['has_sku_data']
        sku_count = result['sku_count']

        self.last_summary = self.calculate_summary_data(result['tracker_df'], marketplace)

        success_msg = f"Report created successfully at:\n{output_file}"
        if marketplace == "Swiggy":
             output_folder = result.get('output_folder')
             success_msg = f"Main report created: {output_file.name}\n\nIndividual PO workbooks created in:\n{output_folder.name}"

        sku_status = f"SKU Data: ✅ Available ({sku_count} SKUs)" if has_sku_data else f"SKU Data: ❌ Not Available ({marketplace} format)"
        summary_text = (
            f"📊 SUMMARY REPORT 📊\n\n"
            f"Total POs: {self.last_summary['total_pos']}\n"
            f"Order Date Range: {self.last_summary['min_date']} to {self.last_summary['max_date']}\n"
            f"Total Units Ordered: {format_indian(self.last_summary['total_units'])}\n"
            f"Total Order Value: ₹ {format_indian(self.last_summary['total_value'])}\n"
            f"{sku_status}"
        )

        SmoothDialog.show_info(self, "Success", success_msg)
        SmoothDialog.show_info(self, "PO Summary", summary_text)

        # Send email prompt
        wants_email = SmoothDialog.ask_question(
            self, 'Send Email', 'Do you want to send this summary via email?'
        )

        if wants_email:
            self.status_label.setText("Sending email...")
            QApplication.processEvents() # Force UI update before blocking email send (or ideally move this to a thread too)

            success, err_msg = EmailService.send_email_summary(
                marketplace,
                self.last_summary,
                result['tracker_df'],
                result['sku_df']
            )
            self.status_label.setText("")

            if success:
                recipient_info = f"To: {Config.DEFAULT_RECIPIENT}"
                if Config.CC_RECIPIENTS:
                    cc_list = "\n".join([f"  • {email}" for email in Config.CC_RECIPIENTS])
                    recipient_info += f"\n\nCC:\n{cc_list}"
                SmoothDialog.show_info(self, "Email Sent", f"Summary sent successfully to:\n\n{recipient_info}")
            else:
                SmoothDialog.show_error(self, "Email Error", f"Failed to send email:\n{err_msg}")

        # Ask to open files
        if marketplace == "Swiggy":
            if SmoothDialog.ask_question(self, "Open File", "Do you want to open the main Excel report?"):
                os.startfile(output_file)
            if SmoothDialog.ask_question(self, "Open Folder", "Do you want to open the PO workbooks folder?"):
                os.startfile(result['output_folder'])
        else:
            if SmoothDialog.ask_question(self, "Open File", "Do you want to open the generated Excel file?"):
                os.startfile(output_file)
                self.close()

def main():
    app = QApplication(sys.argv)

    window = POReportApp()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
