#!/usr/bin/env python3
"""
Headless GUI test for CI/CD pipeline.
Tests that the GUI components can be instantiated without display errors.
"""

import sys
import os

# Add project root to path
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, project_root)

def test_gui_headless():
    """Test GUI initialization in headless environment"""
    print("Starting headless GUI test...")

    try:
        # Set up headless environment for PyQt6
        os.environ['QT_QPA_PLATFORM'] = 'offscreen'

        # Import PyQt6 components
        from PyQt6.QtWidgets import QApplication
        from PyQt6.QtCore import Qt

        print("SUCCESS: PyQt6 imports successful")

        # Create QApplication instance
        app = QApplication([])

        print("SUCCESS: QApplication created successfully")

        # Import and test our GUI classes
        from gui import MainWindow, ProcessingThread

        print("SUCCESS: GUI module imports successful")

        # Test MainWindow instantiation
        main_window = MainWindow()

        print("SUCCESS: MainWindow instantiated successfully")

        # Test that key UI elements exist
        assert hasattr(main_window, 'csv_path_edit'), "CSV path input not found"
        assert hasattr(main_window, 'tenant_id_edit'), "Tenant ID input not found"
        assert hasattr(main_window, 'client_id_edit'), "Client ID input not found"
        assert hasattr(main_window, 'client_secret_edit'), "Client secret input not found"
        assert hasattr(main_window, 'start_btn'), "Start button not found"
        assert hasattr(main_window, 'progress_bar'), "Progress bar not found"
        assert hasattr(main_window, 'log_text'), "Log output not found"

        print("SUCCESS: All required UI elements found")

        # Test ProcessingThread can be instantiated
        thread = ProcessingThread("test.csv", ".", "test_tenant", "test_client", "test_secret")
        assert thread is not None, "ProcessingThread could not be instantiated"

        print("SUCCESS: ProcessingThread instantiated successfully")

        # Test window properties
        assert main_window.windowTitle() == "SharePoint Permissions Excel Tool", "Window title incorrect"

        print("SUCCESS: Window properties correct")

        # Clean up
        main_window.close()
        app.quit()

        print("SUCCESS: GUI cleanup successful")
        print("\nSUCCESS: All GUI tests passed!")

        return True

    except ImportError as e:
        print(f"ERROR: Import error: {e}")
        return False
    except Exception as e:
        print(f"ERROR: GUI test failed: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_gui_components():
    """Test individual GUI components"""
    print("\nTesting individual GUI components...")

    try:
        os.environ['QT_QPA_PLATFORM'] = 'offscreen'

        from PyQt6.QtWidgets import (
            QApplication, QMainWindow, QWidget, QVBoxLayout,
            QHBoxLayout, QLabel, QLineEdit, QPushButton,
            QTextEdit, QProgressBar, QGroupBox, QCheckBox
        )
        from PyQt6.QtCore import QThread, pyqtSignal

        app = QApplication([])

        # Test basic widget creation
        widgets = [
            QMainWindow(),
            QWidget(),
            QVBoxLayout(),
            QHBoxLayout(),
            QLabel("Test"),
            QLineEdit(),
            QPushButton("Test"),
            QTextEdit(),
            QProgressBar(),
            QGroupBox("Test"),
            QCheckBox("Test")
        ]

        print("SUCCESS: All PyQt6 widgets can be instantiated")

        # Test QThread subclassing (like our ProcessingThread)
        class TestThread(QThread):
            test_signal = pyqtSignal(str)

            def run(self):
                self.test_signal.emit("test")

        test_thread = TestThread()
        assert test_thread is not None

        print("SUCCESS: QThread subclassing works correctly")

        app.quit()
        print("SUCCESS: Component tests completed successfully")

        return True

    except Exception as e:
        print(f"ERROR: Component test failed: {e}")
        return False

def main():
    """Run all GUI tests"""
    print("=" * 50)
    print("SharePoint Permissions Exceler - GUI Tests")
    print("=" * 50)

    # Test 1: Basic GUI functionality
    test1_passed = test_gui_headless()

    # Test 2: Individual components
    test2_passed = test_gui_components()

    # Summary
    print("\n" + "=" * 50)
    print("TEST SUMMARY")
    print("=" * 50)
    print(f"GUI Initialization Test: {'PASSED' if test1_passed else 'FAILED'}")
    print(f"Component Tests: {'PASSED' if test2_passed else 'FAILED'}")

    if test1_passed and test2_passed:
        print("\nALL TESTS PASSED!")
        sys.exit(0)
    else:
        print("\nSOME TESTS FAILED!")
        sys.exit(1)

if __name__ == "__main__":
    main()