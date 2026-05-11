import sys
import time
import json
import socket
import time

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QLineEdit, QPushButton, 
                             QMessageBox, QScrollArea, QFileDialog, QComboBox)
from PyQt6.QtCore import Qt, QThread, pyqtSignal

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options


class AutomationThread(QThread):
    """Thread for running automate_device"""
    finished = pyqtSignal(str, bool, str)  # serial_number, success, message
    progress = pyqtSignal(str)  # progress message
    # bootloader_status = pyqtSignal(int, bool)  # True = success, False = fail
    serial_number_verify_status = pyqtSignal(int,str)
    firmware_verify_status = pyqtSignal(int, bool)  # True = match, False = mismatch


    def __init__(self, firmware_version,  firmware_path,  driver_path, chromefortestbinary_path, cycle_number, serial_port, row_index):
        super().__init__()
        self.firmware_version = firmware_version
        self.firmware_path = firmware_path
        self.driver_path = driver_path
        self.chromefortestbinary_path = chromefortestbinary_path
        self.cycle_number = cycle_number
        self.serial_port = serial_port
        self.row_index = row_index

    
    def run(self):
        """Run the automation in a separate thread"""
        try:
            # DON'T send serial data here anymore - it's now handled inside automate_device
            self.progress.emit(f"Starting automation for DUT {self.cycle_number}...")
            
            # Run automation (serial commands are now sent inside this function)
            automate_device(
                firmware_version=self.firmware_version,
                firmware_path=self.firmware_path,
                driver_path=self.driver_path,
                chromefortestbinary_path=self.chromefortestbinary_path,
                serial_number_verify_callback=lambda serial_number:self.serial_number_verify_status.emit(self.row_index,serial_number),
                firmware_verify_callback=lambda ok: self.firmware_verify_status.emit(self.row_index, ok),
                serial_port=self.serial_port,  # Pass serial_port
                cycle_number=self.cycle_number  # Pass cycle_number
            )
            
            self.finished.emit(f"DUT {self.cycle_number}", True, "Successfully processed")
            
        except Exception as e:
            self.finished.emit(self.cycle_number, False, str(e))

class SerialNumberApp(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # Initialize serial port as None
        self.serial_port = None
        self.current_thread = None
        self.automation_queue = []
        self.is_processing = False

        self.serial_inputs = []
        self.serial_verify_indicators = []

        
        self.setWindowTitle("Serial Number Input")
        self.setGeometry(100, 100, 700, 750)
        
        # Store single record in memory
        self.saved_data = None
        
        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # Title
        title = QLabel("Serial Number Manager")
        title.setStyleSheet("font-size: 18px; font-weight: bold; padding: 10px;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(title)

 
        # File path selection section
        path_section = QWidget()
        path_layout = QVBoxLayout(path_section)
        path_layout.setContentsMargins(10, 10, 10, 10)
        
        # Firmware path
        firmware_layout = QHBoxLayout()
        firmware_label = QLabel("Firmware:")
        firmware_label.setMinimumWidth(100)

        self.firmware_path = QLineEdit()
        self.firmware_path.setPlaceholderText("Select firmware file...")
        self.firmware_path.setReadOnly(True)
        
        firmware_btn = QPushButton("Browse")
        firmware_btn.setMaximumWidth(80)
        firmware_btn.clicked.connect(self.select_firmware)

        firmware_layout.addWidget(firmware_label)
        firmware_layout.addWidget(self.firmware_path)
        firmware_layout.addWidget(firmware_btn)

        # Firmware version input
        firmware_version_layout = QHBoxLayout()
        firmware_version_label = QLabel("Firmware version:")
        firmware_version_label.setMinimumWidth(100)

        self.input_firmware_version = QLineEdit()
        self.input_firmware_version.setPlaceholderText("Enter firmware version...")
        self.input_firmware_version.setStyleSheet("padding: 5px; font-size: 12px;")
        
        host_id_label = QLabel("Host ID:")
        host_id_label.setMinimumWidth(50)

        self.input_host_id_low = QLineEdit()
        self.input_host_id_low.setPlaceholderText("First host id...")
        self.input_host_id_low.setStyleSheet("padding: 5px; font-size: 12px;")

        self.input_host_id_high = QLineEdit()
        self.input_host_id_high.setPlaceholderText("Last host id...")
        self.input_host_id_high.setStyleSheet("padding: 5px; font-size: 12px;")

        firmware_version_layout.addWidget(firmware_version_label)
        firmware_version_layout.addWidget(self.input_firmware_version)
        firmware_version_layout.addWidget(host_id_label)
        firmware_version_layout.addWidget(self.input_host_id_low)
        firmware_version_layout.addWidget(self.input_host_id_high)


        path_layout.addLayout(firmware_layout)
        path_layout.addLayout(firmware_version_layout)
        
        path_section.setStyleSheet("""
            QWidget {
                background-color: #f5f5f5;
                border-radius: 5px;
            }
            QLineEdit {
                padding: 5px;
                background-color: white;
                border: 1px solid #ddd;
            }
            QPushButton {
                padding: 5px 10px;
                background-color: #2196F3;
                color: white;
                border: none;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #0b7dda;
            }
        """)
        
        main_layout.addWidget(path_section)
        
        # Create scroll area for input fields
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        
        # Create 8 input fields with buttons
        self.serial_inputs = []
 
        self.firmware_buttons = []    # Track  firmware buttons
        self.last_clicked_button = None  # Track the last clicked button
                
        for i in range(1, 9):
            field_layout = QHBoxLayout()
            
            label = QLabel(f"Serial number {i}:")
            label.setMinimumWidth(80)
            label.setFixedWidth(100)
            label.setStyleSheet("font-size: 12px;")
            
            # Display box instead of input field
            # display_box = QLabel()
            # display_box.setMinimumWidth(200)
            # display_box.setStyleSheet("""
            #     font-size: 12px;
            #     padding: 5px;
            #     border: 1px solid #ccc;
            #     border-radius: 3px;
            #     background-color: #f5f5f5;
            # """)

            field_layout.addWidget(label)
            # field_layout.addWidget(display_box)
            
            # Serial verify indicator
            sn_led = QLabel()
            sn_led.setFixedSize(18, 18)
            sn_led.setStyleSheet("""
                QLabel {
                    background-color: #999;
                    border-radius: 9px;
                    border: 1px solid #666;
                }
            """)
            sn_led.setAlignment(Qt.AlignmentFlag.AlignCenter)

            field_layout.addWidget(sn_led)

            scroll_layout.addLayout(field_layout)

            # self.serial_inputs.append(display_box)
            self.serial_verify_indicators.append(sn_led)

        
        scroll.setWidget(scroll_content)
        main_layout.addWidget(scroll)

        # Button layout
        button_layout = QHBoxLayout()
        
        # Upload button
        save_btn = QPushButton("Upload")
        save_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                padding: 10px;
                font-size: 14px;
                border: none;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        save_btn.clicked.connect(self.save_serial_numbers)
        
        # Clear button (moved before Reset)
        clear_btn = QPushButton("Clear")
        clear_btn.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                padding: 10px;
                font-size: 14px;
                border: none;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #da190b;
            }
        """)
        clear_btn.clicked.connect(self.clear_fields)

        button_layout.addWidget(save_btn)
        button_layout.addWidget(clear_btn)
        main_layout.addLayout(button_layout)
        
        # Status label
        self.status_label = QLabel("No data saved")
        self.status_label.setStyleSheet("font-size: 11px; color: #666; padding: 5px;")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(self.status_label)
    
    
    def save_serial_numbers(self):
        firmware_version = self.input_firmware_version.text().strip()

        if not firmware_version:
            QMessageBox.warning(self, "No input firmware version", "Please enter firmware version!")
            return

        host_id_low = self.input_host_id_low.text().strip()
        host_id_high = self.input_host_id_high.text().strip()
      
        if not host_id_low.isdigit() or not host_id_high.isdigit():
            QMessageBox.warning(self, "Invalid Input", "Please enter correct host ID number!")
            return

        host_id_low = int(host_id_low)
        host_id_high = int(host_id_high)

        if host_id_high < host_id_low:
            QMessageBox.warning(self, "Invalid Input", "Host ID high must be greater than or equal to Host ID low!")
            return

        total_dut = (host_id_high - host_id_low) + 1

        self.saved_data = {
            "firmware_version": firmware_version,
            "host_id_low":host_id_low,
            "host_id_high":host_id_high,
            "total_dut": total_dut, 
        }
        
        print(f"Host id low = {self.saved_data['host_id_low']}")
        print(f"Host id high = {self.saved_data['host_id_high']}")
        print(f"Total DUT: {self.saved_data['total_dut']}")

        QMessageBox.information(self, "Begin upload", f"Firmware version saved: {firmware_version}\nHost id low: {host_id_low}\nHost id high: {host_id_high}\nTotal number of DUT: {total_dut}")
        
        self.upload_package()
    
    def clear_fields(self):
        reply = QMessageBox.question(
            self, 
            "Clear Fields", 
            "Are you sure you want to clear all input fields?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            for input_field in self.serial_inputs:
                input_field.clear()
                for row_idx, _ in enumerate(self.serial_verify_indicators):
                    self.serial_verify_indicators[row_idx].setStyleSheet("""
                                                                        background-color: #999;
                                                                        border-radius: 9px;
                                                                        border: 1px solid #666;
                                                                        """
                                                                        )
  
    def select_firmware(self):
        """Open file dialog to select firmware file"""
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Select Firmware File",
            "",
            "Firmware Files (*.acfr);;All Files (*)"
        )
        
        if filename:
            self.firmware_path.setText(filename)
            print(f"Firmware selected: {filename}")
    
    def upload_package(self):
        """Prepare and start automation queue"""
        print("=" * 50)
        print("Uploading package")
        print(f"Firmware version: {self.saved_data['firmware_version']}")
        print(f"Host id low: {self.saved_data['host_id_low']}")
        print(f"Host id high: {self.saved_data['host_id_high']}")
        print(f"Total Entries: {self.saved_data['total_dut']}")
        print("=" * 50)
        
        # Get file paths
        firmware = self.firmware_path.text()
        total_dut = int(self.saved_data['total_dut'])

        # Validate that paths are selected
        if not firmware:
            QMessageBox.warning(self, "Missing File", "Please select a firmware file!")
            return
        
        # Build automation queue
        self.automation_queue = []
        cycle_number = 0x01
        
        for i in range(1, total_dut + 1):  # ← loop by count, not serial numbers
            self.automation_queue.append({
                'cycle_number': cycle_number,
                'firmware': firmware,
                'firmware_version': self.saved_data['firmware_version'],
                'serial_number': None  # ← will be filled later by API
            })
            cycle_number += 1

        self.process_next_in_queue()

    def update_serial_number_verify_status(self,row,serial_number):
        self.serial_inputs[row].setText(serial_number)
        
    def update_firmware_verify_status(self, row, success):
        led = self.serial_verify_indicators[row]
        if success:
            led.setStyleSheet("""
                QLabel {
                    background-color: #4CAF50;
                    border-radius: 9px;
                    border: 2px solid #2E7D32;
                }
            """)
        else:
            led.setStyleSheet("""
                QLabel {
                    background-color: #f44336;
                    border-radius: 9px;
                    border: 2px solid #C62828;
                }
            """)

    def process_next_in_queue(self):
        if not self.automation_queue:
            print("\n" + "=" * 50)
            print("All automation tasks completed!")
            print("=" * 50)
            self.status_label.setText("✓ All tasks completed")
            self.status_label.setStyleSheet("font-size: 11px; color: #4CAF50; padding: 5px;")
            self.is_processing = False
            return
        
        # Get next task
        task = self.automation_queue.pop(0)
        self.is_processing = True
        
        # Reset indicators for this row
        row_idx = task['cycle_number'] - 1
        self.serial_verify_indicators[row_idx].setStyleSheet("""
                background-color: #999;
                border-radius: 9px;
                border: 1px solid #666;
            """)

        print(f"\nProcessing DUT{task['cycle_number']}: (Cycle 0x{task['cycle_number']:02X})")
        
        # Update status
        self.status_label.setText(f"⏳ Processing DUT {task['cycle_number']}...")
        self.status_label.setStyleSheet("font-size: 11px; color: #2196F3; padding: 5px;")
        
        # Create and start thread
        self.current_thread = AutomationThread(
            firmware_version=task['firmware_version'],
            firmware_path=task['firmware'],
            driver_path=r"D:\MULTIPROGRAMMER\chromedriver-win64\chromedriver-win64\chromedriver.exe",
            chromefortestbinary_path = r"D:\MULTIPROGRAMMER\chrome-win64\chrome-win64\chrome.exe",
            cycle_number=task['cycle_number'],
            serial_port=self.serial_port,
            row_index=task['cycle_number'] - 1 
        )
        
        # Connect signals
        self.current_thread.progress.connect(self.on_automation_progress)
        self.current_thread.finished.connect(self.on_automation_finished)
        self.current_thread.serial_number_verify_status.connect(self.update_serial_number_verify_status)
        self.current_thread.firmware_verify_status.connect(self.update_firmware_verify_status)
        
        # Start thread
        self.current_thread.start()
    
    def on_automation_progress(self, message):
        """Handle progress updates from automation thread"""
        print(message)
    
    def on_automation_finished(self, serial_number, success, message):
        """Handle automation thread completion"""
        if success:
            print(f"{serial_number} Successfully processed")
            self.status_label.setText(f"Completed {serial_number}")
            self.status_label.setStyleSheet("font-size: 11px; color: #4CAF50; padding: 5px;")
        else:
            print(f"Error processing {serial_number}: {message}")
            self.status_label.setText(f"Failed {serial_number}")
            self.status_label.setStyleSheet("font-size: 11px; color: #f44336; padding: 5px;")
            QMessageBox.warning(
                self,
                "Processing Error",
                f"Failed to process {serial_number}\n\nError: {message}"
            )
        
        # Process next item in queue
        self.process_next_in_queue()
    
    def closeEvent(self, event):
        """Handle window close event"""
        # Wait for current thread to finish
        if self.current_thread and self.current_thread.isRunning():
            reply = QMessageBox.question(
                self,
                "Task Running",
                "An automation task is currently running. Are you sure you want to exit?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            
            if reply == QMessageBox.StandardButton.No:
                event.ignore()
                return
            
            self.current_thread.terminate()
            self.current_thread.wait()
        
        # Close serial port
        if self.serial_port and self.serial_port.is_open:
            self.serial_port.close()
            print("Serial port closed on exit")
        
        event.accept()


def wait_for_device_ready(ip="192.168.0.100", port=80, timeout=30):
    """
    Wait for the device's web server to be ready by attempting TCP connection
    Returns True if device is ready, False if timeout
    """
    start_time = time.time()
    
    while (time.time() - start_time) < timeout:
        try:
            # Try to establish TCP connection to the web server
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(1)
            result = sock.connect_ex((ip, port))
            sock.close()
            
            if result == 0:
                return True
        except:
            pass

        time.sleep(0.5)  # Check every 0.5 seconds
    return False


def automate_device(
        firmware_version, 
        firmware_path, 
        driver_path,
        chromefortestbinary_path,
        serial_number_verify_callback,
        firmware_verify_callback,
        serial_port,cycle_number
        ):
    driver = None

    try:
        data_bytes_before_firmware = bytes([0x41, 0x01, 0xFF, 0x0D])
        serial_port.write(data_bytes_before_firmware)
        time.sleep(2)

        data_service = [0x41, 0x01, cycle_number+8, 0x0D]
        data_bytes_service = bytes(data_service)
        serial_port.write(data_bytes_service)
        
        # CRITICAL FIX: Wait for device web server to actually be ready
        if not wait_for_device_ready(ip="192.168.0.100", port=80, timeout=30):
            raise Exception("Device web server did not become ready in time")

        # Now that device is confirmed ready, start browser
        options = Options()
        options.binary_location = chromefortestbinary_path
        options.page_load_strategy = 'eager'
        options.add_argument('--disable-gpu')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-extensions')

        service = Service(driver_path)
        driver = webdriver.Chrome(service=service, options=options)

        # Navigate to factory config page
        driver.get("http://192.168.0.100/#/login")
        
        # Wait for page to be fully loaded
        WebDriverWait(driver, 5).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )

        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input[aria-label="Username"]'))
        )
        
        driver.find_element(By.CSS_SELECTOR, 'input[aria-label="Username"]').send_keys("admin")
        driver.find_element(By.CSS_SELECTOR, 'input[aria-label="Password"]').send_keys("admin")
        time.sleep(0.5)
        driver.find_element(By.XPATH, '//span[text()="Login"]').click()
        time.sleep(0.5)
        driver.find_element(By.XPATH, '//div[text()="System"]').click()
        time.sleep(0.5)
        file_input = driver.find_element(By.CSS_SELECTOR, 'input[type="file"]')
        file_input.send_keys(firmware_path)  
        time.sleep(0.5)
        driver.find_element(By.XPATH, '//span[text()="Upload"]').click()
        time.sleep(0.5)

        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input[aria-label="Username"]'))
        )
        
        driver.find_element(By.CSS_SELECTOR, 'input[aria-label="Username"]').send_keys("admin")
        driver.find_element(By.CSS_SELECTOR, 'input[aria-label="Password"]').send_keys("admin")
        time.sleep(0.5)
        driver.find_element(By.XPATH, '//span[text()="Login"]').click()
        time.sleep(3)
        try:
            driver.get("http://192.168.0.100/config.json")
            time.sleep(0.5)
            
            json_text = driver.find_element(By.TAG_NAME, "body").text
            data = json.loads(json_text)
            serial_number_from_device = data["deviceInfo"]["serialNumber"]
            firmware_from_device = data["deviceInfo"]["firmwareVersion"]

            print(f"Serial Number from device: {serial_number_from_device}")
            print(f"Firmware verion from device: {firmware_from_device}")


            serial_number_verify_callback(serial_number_from_device)
            fw_match = (firmware_from_device == firmware_version)
            firmware_verify_callback(fw_match)
            
        except Exception as e:
            print(f"Failed to verify serial number: {e}")
            firmware_verify_callback(False)
        

    except Exception as e:
        print(f"Automation error: {e}")
        firmware_verify_callback(False)
    
    finally:
        if driver is not None:
            try:
                driver.quit()
            except:
                pass



def main():
    app = QApplication(sys.argv)
    window = SerialNumberApp()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()