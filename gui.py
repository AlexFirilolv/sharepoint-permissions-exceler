import sys
import os
from pathlib import Path
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QFileDialog, QTextEdit,
    QProgressBar, QGroupBox, QCheckBox, QMessageBox
)
from PyQt6.QtCore import QThread, pyqtSignal, Qt
from PyQt6.QtGui import QFont, QIcon
import pandas as pd
from main import (
    GraphAPIClient,
    create_excel_report_with_tables,
    validate_azure_credentials
)
from dotenv import load_dotenv, set_key, find_dotenv

class ProcessingThread(QThread):
    """
    Worker thread to process SharePoint permissions without blocking the GUI.
    """
    progress = pyqtSignal(int)  # Progress percentage
    status = pyqtSignal(str)  # Status message
    log = pyqtSignal(str)  # Log message
    finished = pyqtSignal(bool, str)  # Success flag and message

    def __init__(self, csv_path, output_dir, tenant_id, client_id, client_secret):
        super().__init__()
        self.csv_path = csv_path
        self.output_dir = output_dir
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self._is_cancelled = False

    def cancel(self):
        """Request cancellation of the processing."""
        self._is_cancelled = True

    def run(self):
        """Execute the processing logic."""
        try:
            # Initialize Graph API client
            self.status.emit("Initializing Azure credentials...")
            self.progress.emit(5)

            if self.tenant_id and self.client_id and self.client_secret:
                self.log.emit("Authenticating with Microsoft Graph API...")
                graph_client = GraphAPIClient(self.tenant_id, self.client_id, self.client_secret)

                if graph_client.enabled:
                    self.log.emit("✓ Successfully authenticated with Azure!")
                else:
                    self.log.emit(f"✗ Authentication failed: {graph_client.get_auth_error()}")
                    self.log.emit("Continuing without security group expansion...")
                    graph_client = GraphAPIClient()
            else:
                self.log.emit("No Azure credentials provided - skipping security group expansion")
                graph_client = GraphAPIClient()

            if self._is_cancelled:
                self.finished.emit(False, "Processing cancelled")
                return

            # Read CSV file
            self.status.emit("Reading CSV file...")
            self.progress.emit(15)
            self.log.emit(f"Loading CSV file: {self.csv_path}")

            df = pd.read_csv(self.csv_path)
            self.log.emit(f"✓ Loaded {len(df)} permission entries")

            if self._is_cancelled:
                self.finished.emit(False, "Processing cancelled")
                return

            # Generate Parent Permissions Report
            self.status.emit("Generating Parent Permissions Report...")
            self.progress.emit(30)
            self.log.emit("\n--- Parent Permissions Report ---")

            parent_df = df[df['Item Type'].isin(['Web', 'List'])].copy()
            parent_df = parent_df.drop(columns=['Link ID', 'Link Type', 'AccessViaLinkID'], errors='ignore')
            parent_groups = parent_df.groupby('Resource Path')

            self.log.emit(f"Processing {len(parent_groups)} parent resources...")

            parent_output_path = os.path.join(self.output_dir, 'Parent_Permissions.xlsx')
            create_excel_report_with_tables(parent_groups, parent_output_path, graph_client)
            self.log.emit(f"✓ Created: {parent_output_path}")

            if self._is_cancelled:
                self.finished.emit(False, "Processing cancelled")
                return

            # Generate Unique Permissions Report
            self.status.emit("Generating Unique Permissions Report...")
            self.progress.emit(60)
            self.log.emit("\n--- Unique Permissions Report ---")

            unique_df = df[~df['Item Type'].isin(['Web', 'List'])].copy()

            if not unique_df.empty:
                parent_paths = parent_df['Resource Path'].unique()

                def find_parent_library(path):
                    for parent in parent_paths:
                        if path.startswith(parent):
                            return parent
                    return '/'.join(path.split('/')[:3])

                unique_df['Document Library'] = unique_df['Resource Path'].apply(find_parent_library)
                unique_groups = unique_df.groupby('Document Library')

                self.log.emit(f"Processing {len(unique_groups)} document libraries with unique permissions...")

                unique_output_path = os.path.join(self.output_dir, 'Unique_Permissions.xlsx')
                create_excel_report_with_tables(unique_groups, unique_output_path, graph_client)
                self.log.emit(f"✓ Created: {unique_output_path}")
            else:
                self.log.emit("No unique permissions found in the data")

            if self._is_cancelled:
                self.finished.emit(False, "Processing cancelled")
                return

            # Complete
            self.status.emit("Complete!")
            self.progress.emit(100)
            self.log.emit("\n" + "=" * 50)
            self.log.emit("Report generation complete!")
            self.log.emit("=" * 50)

            self.finished.emit(True, "Reports generated successfully!")

        except Exception as e:
            self.log.emit(f"\n✗ Error: {str(e)}")
            import traceback
            self.log.emit(traceback.format_exc())
            self.finished.emit(False, f"Error: {str(e)}")


class MainWindow(QMainWindow):
    """
    Main application window for SharePoint Permissions Excel Tool.
    """
    def __init__(self):
        super().__init__()
        self.processing_thread = None
        self.init_ui()
        self.load_env_credentials()

    def init_ui(self):
        """Initialize the user interface."""
        self.setWindowTitle("SharePoint Permissions Excel Tool")
        self.setMinimumSize(800, 700)

        # Main widget and layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        layout.setSpacing(15)

        # Title
        title = QLabel("SharePoint Permissions Report Generator")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        # File Selection Group
        file_group = QGroupBox("1. Input File")
        file_layout = QVBoxLayout()

        csv_layout = QHBoxLayout()
        self.csv_path_edit = QLineEdit()
        self.csv_path_edit.setPlaceholderText("Select SharePoint permissions CSV file...")
        csv_browse_btn = QPushButton("Browse...")
        csv_browse_btn.clicked.connect(self.browse_csv_file)
        csv_layout.addWidget(QLabel("CSV File:"))
        csv_layout.addWidget(self.csv_path_edit, stretch=1)
        csv_layout.addWidget(csv_browse_btn)
        file_layout.addLayout(csv_layout)

        output_layout = QHBoxLayout()
        self.output_dir_edit = QLineEdit()
        self.output_dir_edit.setPlaceholderText("Select output directory for Excel reports...")
        output_browse_btn = QPushButton("Browse...")
        output_browse_btn.clicked.connect(self.browse_output_dir)
        output_layout.addWidget(QLabel("Output Dir:"))
        output_layout.addWidget(self.output_dir_edit, stretch=1)
        output_layout.addWidget(output_browse_btn)
        file_layout.addLayout(output_layout)

        file_group.setLayout(file_layout)
        layout.addWidget(file_group)

        # Azure Credentials Group
        creds_group = QGroupBox("2. Azure AD Credentials (Optional - for Security Group Expansion)")
        creds_layout = QVBoxLayout()

        # Load from .env checkbox
        self.load_env_checkbox = QCheckBox("Load credentials from .env file")
        self.load_env_checkbox.setChecked(True)
        self.load_env_checkbox.stateChanged.connect(self.toggle_credential_fields)
        creds_layout.addWidget(self.load_env_checkbox)

        # Tenant ID
        tenant_layout = QHBoxLayout()
        tenant_layout.addWidget(QLabel("Tenant ID:"))
        self.tenant_id_edit = QLineEdit()
        self.tenant_id_edit.setPlaceholderText("xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx")
        tenant_layout.addWidget(self.tenant_id_edit)
        creds_layout.addLayout(tenant_layout)

        # Client ID
        client_layout = QHBoxLayout()
        client_layout.addWidget(QLabel("Client ID:"))
        self.client_id_edit = QLineEdit()
        self.client_id_edit.setPlaceholderText("xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx")
        client_layout.addWidget(self.client_id_edit)
        creds_layout.addLayout(client_layout)

        # Client Secret
        secret_layout = QHBoxLayout()
        secret_layout.addWidget(QLabel("Client Secret:"))
        self.client_secret_edit = QLineEdit()
        self.client_secret_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.client_secret_edit.setPlaceholderText("Enter client secret...")
        secret_layout.addWidget(self.client_secret_edit)
        creds_layout.addLayout(secret_layout)

        # Save credentials checkbox
        self.save_creds_checkbox = QCheckBox("Save credentials to .env file")
        creds_layout.addWidget(self.save_creds_checkbox)

        # Help button
        help_btn = QPushButton("How to get Azure credentials?")
        help_btn.clicked.connect(self.show_azure_help)
        creds_layout.addWidget(help_btn)

        creds_group.setLayout(creds_layout)
        layout.addWidget(creds_group)

        # Progress Section
        progress_group = QGroupBox("3. Processing")
        progress_layout = QVBoxLayout()

        self.status_label = QLabel("Ready")
        progress_layout.addWidget(self.status_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        progress_layout.addWidget(self.progress_bar)

        progress_group.setLayout(progress_layout)
        layout.addWidget(progress_group)

        # Log Output
        log_group = QGroupBox("4. Log Output")
        log_layout = QVBoxLayout()

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(200)
        log_layout.addWidget(self.log_text)

        log_group.setLayout(log_layout)
        layout.addWidget(log_group)

        # Control Buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        self.start_btn = QPushButton("Generate Reports")
        self.start_btn.setMinimumWidth(150)
        self.start_btn.clicked.connect(self.start_processing)
        button_layout.addWidget(self.start_btn)

        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.setMinimumWidth(150)
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.clicked.connect(self.cancel_processing)
        button_layout.addWidget(self.cancel_btn)

        button_layout.addStretch()
        layout.addLayout(button_layout)

        # Set default output directory to current directory
        self.output_dir_edit.setText(os.getcwd())

    def load_env_credentials(self):
        """Load credentials from .env file if available."""
        load_dotenv()

        tenant_id = os.getenv('AZURE_TENANT_ID', '')
        client_id = os.getenv('AZURE_CLIENT_ID', '')
        client_secret = os.getenv('AZURE_CLIENT_SECRET', '')

        if tenant_id and tenant_id != 'your-tenant-id-here':
            self.tenant_id_edit.setText(tenant_id)
        if client_id and client_id != 'your-client-id-here':
            self.client_id_edit.setText(client_id)
        if client_secret and client_secret != 'your-client-secret-here':
            self.client_secret_edit.setText(client_secret)

    def toggle_credential_fields(self):
        """Enable/disable credential fields based on checkbox state."""
        enabled = not self.load_env_checkbox.isChecked()
        self.tenant_id_edit.setEnabled(enabled)
        self.client_id_edit.setEnabled(enabled)
        self.client_secret_edit.setEnabled(enabled)

        if not enabled:
            # Reload credentials from .env when checkbox is checked
            self.load_env_credentials()

    def browse_csv_file(self):
        """Open file dialog to select CSV file."""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select SharePoint Permissions CSV File",
            os.getcwd(),
            "CSV Files (*.csv);;All Files (*)"
        )
        if file_path:
            self.csv_path_edit.setText(file_path)

    def browse_output_dir(self):
        """Open directory dialog to select output directory."""
        dir_path = QFileDialog.getExistingDirectory(
            self,
            "Select Output Directory",
            self.output_dir_edit.text() or os.getcwd()
        )
        if dir_path:
            self.output_dir_edit.setText(dir_path)

    def show_azure_help(self):
        """Show help dialog for obtaining Azure credentials."""
        help_text = """
<h3>How to Obtain Azure AD Credentials</h3>

<p><b>1.</b> Go to <a href="https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps">Azure Portal - App Registrations</a></p>

<p><b>2.</b> Click "+ New registration"</p>

<p><b>3.</b> After creating the app, note the:
<ul>
<li>Directory (tenant) ID</li>
<li>Application (client) ID</li>
</ul>
</p>

<p><b>4.</b> Go to "Certificates & secrets" > "+ New client secret"</p>

<p><b>5.</b> Copy the secret VALUE immediately (shown only once)</p>

<p><b>6.</b> Go to "API permissions" > "+ Add a permission"</p>

<p><b>7.</b> Select "Microsoft Graph" > "Application permissions"</p>

<p><b>8.</b> Add these permissions:
<ul>
<li>GroupMember.Read.All</li>
<li>User.Read.All</li>
</ul>
</p>

<p><b>9.</b> Click "Grant admin consent"</p>

<p><i>Note: Without these credentials, the tool will still work but will not expand security group members.</i></p>
"""
        msg = QMessageBox()
        msg.setWindowTitle("Azure Credentials Help")
        msg.setTextFormat(Qt.TextFormat.RichText)
        msg.setText(help_text)
        msg.exec()

    def validate_inputs(self):
        """Validate user inputs before processing."""
        csv_path = self.csv_path_edit.text().strip()
        output_dir = self.output_dir_edit.text().strip()

        if not csv_path:
            QMessageBox.warning(self, "Validation Error", "Please select a CSV file.")
            return False

        if not os.path.exists(csv_path):
            QMessageBox.warning(self, "Validation Error", f"CSV file not found:\n{csv_path}")
            return False

        if not output_dir:
            QMessageBox.warning(self, "Validation Error", "Please select an output directory.")
            return False

        if not os.path.exists(output_dir):
            QMessageBox.warning(self, "Validation Error", f"Output directory not found:\n{output_dir}")
            return False

        # Validate Azure credentials if provided
        tenant_id = self.tenant_id_edit.text().strip()
        client_id = self.client_id_edit.text().strip()
        client_secret = self.client_secret_edit.text().strip()

        # Only validate if any credential is provided
        if tenant_id or client_id or client_secret:
            # Check if all three are provided
            if not all([tenant_id, client_id, client_secret]):
                QMessageBox.warning(
                    self,
                    "Validation Error",
                    "If providing Azure credentials, all three fields (Tenant ID, Client ID, Client Secret) must be filled."
                )
                return False

            # Validate credential format
            is_valid, error_msg = validate_azure_credentials(tenant_id, client_id, client_secret)
            if not is_valid:
                QMessageBox.warning(self, "Validation Error", f"Invalid Azure credentials:\n{error_msg}")
                return False

        return True

    def save_credentials_to_env(self):
        """Save credentials to .env file."""
        try:
            tenant_id = self.tenant_id_edit.text().strip()
            client_id = self.client_id_edit.text().strip()
            client_secret = self.client_secret_edit.text().strip()

            env_file = find_dotenv()
            if not env_file:
                env_file = os.path.join(os.getcwd(), '.env')

            set_key(env_file, 'AZURE_TENANT_ID', tenant_id)
            set_key(env_file, 'AZURE_CLIENT_ID', client_id)
            set_key(env_file, 'AZURE_CLIENT_SECRET', client_secret)

            self.append_log(f"✓ Credentials saved to {env_file}")
        except Exception as e:
            self.append_log(f"✗ Failed to save credentials: {str(e)}")

    def start_processing(self):
        """Start the processing thread."""
        if not self.validate_inputs():
            return

        # Save credentials if requested
        if self.save_creds_checkbox.isChecked():
            self.save_credentials_to_env()

        # Disable start button and enable cancel button
        self.start_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)

        # Clear log and reset progress
        self.log_text.clear()
        self.progress_bar.setValue(0)

        # Get inputs
        csv_path = self.csv_path_edit.text().strip()
        output_dir = self.output_dir_edit.text().strip()
        tenant_id = self.tenant_id_edit.text().strip()
        client_id = self.client_id_edit.text().strip()
        client_secret = self.client_secret_edit.text().strip()

        # Create and start processing thread
        self.processing_thread = ProcessingThread(
            csv_path, output_dir, tenant_id, client_id, client_secret
        )
        self.processing_thread.progress.connect(self.update_progress)
        self.processing_thread.status.connect(self.update_status)
        self.processing_thread.log.connect(self.append_log)
        self.processing_thread.finished.connect(self.processing_finished)

        self.append_log("=" * 50)
        self.append_log("Starting report generation...")
        self.append_log("=" * 50)

        self.processing_thread.start()

    def cancel_processing(self):
        """Cancel the processing thread."""
        if self.processing_thread and self.processing_thread.isRunning():
            self.append_log("\nCancelling processing...")
            self.processing_thread.cancel()
            self.cancel_btn.setEnabled(False)

    def update_progress(self, value):
        """Update progress bar."""
        self.progress_bar.setValue(value)

    def update_status(self, message):
        """Update status label."""
        self.status_label.setText(message)

    def append_log(self, message):
        """Append message to log text area."""
        self.log_text.append(message)
        # Auto-scroll to bottom
        scrollbar = self.log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def processing_finished(self, success, message):
        """Handle processing completion."""
        self.start_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)

        if success:
            QMessageBox.information(self, "Success", message)
        else:
            if "cancelled" not in message.lower():
                QMessageBox.critical(self, "Error", message)


def main():
    """Main entry point for the GUI application."""
    app = QApplication(sys.argv)
    app.setApplicationName("SharePoint Permissions Excel Tool")

    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
