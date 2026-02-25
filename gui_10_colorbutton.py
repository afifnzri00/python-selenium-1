import serial
import serial.tools.list_ports
import sys
import time
import subprocess
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
    bootloader_status = pyqtSignal(int, bool)  # True = success, False = fail
    serial_verify_status = pyqtSignal(int, bool)  # True = match, False = mismatch


    def __init__(self, row_index, serial_number, bootloader_path, firmware_path, bat_file, driver_path, chromefortestbinary_path, cycle_number, serial_port):
        super().__init__()
        self.serial_number = serial_number
        self.bootloader_path = bootloader_path
        self.firmware_path = firmware_path
        self.bat_file = bat_file
        self.driver_path = driver_path
        self.chromefortestbinary_path = chromefortestbinary_path
        self.cycle_number = cycle_number
        self.serial_port = serial_port
        self.row_index = row_index

    
    def run(self):
        """Run the automation in a separate thread"""
        try:
            # DON'T send serial data here anymore - it's now handled inside automate_device
            self.progress.emit(f"Starting automation for {self.serial_number}...")
            
            # Run automation (serial commands are now sent inside this function)
            automate_device(
                serial_number=self.serial_number,
                bootloader_path=self.bootloader_path,
                firmware_path=self.firmware_path,
                bat_file=self.bat_file,
                driver_path=self.driver_path,
                chromefortestbinary_path=self.chromefortestbinary_path,
                bootloader_callback=lambda ok: self.bootloader_status.emit(self.row_index, ok),
                serial_verify_callback=lambda ok: self.serial_verify_status.emit(self.row_index, ok),
                serial_port=self.serial_port,  # Pass serial_port
                cycle_number=self.cycle_number  # Pass cycle_number
            )
            
            self.finished.emit(self.serial_number, True, "Successfully processed")
            
        except Exception as e:
            self.finished.emit(self.serial_number, False, str(e))

class SerialNumberApp(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # Initialize serial port as None
        self.serial_port = None
        self.current_thread = None
        self.automation_queue = []
        self.is_processing = False

        self.bootloader_indicators = []
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
        
        # Serial Port Selection Section
        serial_section = QWidget()
        serial_layout = QHBoxLayout(serial_section)
        serial_layout.setContentsMargins(10, 10, 10, 10)
        
        # COM Port Dropdown
        port_label = QLabel("COM Port:")
        port_label.setMinimumWidth(100)
        self.port_combo = QComboBox()
        self.port_combo.setMinimumWidth(150)
        self.refresh_ports()
        
        # Refresh button
        refresh_btn = QPushButton("🔄")
        refresh_btn.setMaximumWidth(40)
        refresh_btn.setToolTip("Refresh COM ports")
        refresh_btn.clicked.connect(self.refresh_ports)
        
        # Connect button
        self.connect_btn = QPushButton("Connect")
        self.connect_btn.setMaximumWidth(100)
        self.connect_btn.clicked.connect(self.connect_serial)
        
        # Disconnect button
        self.disconnect_btn = QPushButton("Disconnect")
        self.disconnect_btn.setMaximumWidth(100)
        self.disconnect_btn.setEnabled(False)
        self.disconnect_btn.clicked.connect(self.disconnect_serial)
        
        # Connection status label
        self.connection_status = QLabel("● Disconnected")
        self.connection_status.setStyleSheet("color: #f44336; font-weight: bold;")
        
        serial_layout.addWidget(port_label)
        serial_layout.addWidget(self.port_combo)
        serial_layout.addWidget(refresh_btn)
        serial_layout.addWidget(self.connect_btn)
        serial_layout.addWidget(self.disconnect_btn)
        serial_layout.addWidget(self.connection_status)
        serial_layout.addStretch()
        
        serial_section.setStyleSheet("""
            QWidget {
                background-color: #f5f5f5;
                border-radius: 5px;
            }
            QComboBox {
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
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        
        main_layout.addWidget(serial_section)
    
        # File path selection section
        path_section = QWidget()
        path_layout = QVBoxLayout(path_section)
        path_layout.setContentsMargins(10, 10, 10, 10)
        
        # Bootloader path
        bootloader_layout = QHBoxLayout()
        bootloader_label = QLabel("Bootloader:")
        bootloader_label.setMinimumWidth(100)
        self.bootloader_path = QLineEdit()
        self.bootloader_path.setPlaceholderText("Select bootloader file...")
        self.bootloader_path.setReadOnly(True)
        bootloader_btn = QPushButton("Browse")
        bootloader_btn.setMaximumWidth(80)
        bootloader_btn.clicked.connect(self.select_bootloader)
        bootloader_layout.addWidget(bootloader_label)
        bootloader_layout.addWidget(self.bootloader_path)
        bootloader_layout.addWidget(bootloader_btn)
        
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
        
        path_layout.addLayout(bootloader_layout)
        path_layout.addLayout(firmware_layout)
        
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
 
        self.bootloader_buttons = []  # Track bootloader buttons
        self.firmware_buttons = []    # Track firmware buttons
        self.last_clicked_button = None  # Track the last clicked button
                
        for i in range(1, 9):
            field_layout = QHBoxLayout()
            
            label = QLabel(f"Serial number {i}:")
            label.setMinimumWidth(80)
            label.setStyleSheet("font-size: 12px;")
            
            input_field = QLineEdit()
            input_field.setPlaceholderText(f"Enter serial number {i}")
            input_field.setStyleSheet("padding: 5px; font-size: 12px;")
            
            # Add for bootloader button for each field
            bootloader_btn = QPushButton("For bootloader")
            bootloader_btn.setMaximumWidth(90)
            bootloader_btn.setStyleSheet("""
                QPushButton {
                    background-color: #999;
                    color: white;
                    padding: 5px;
                    font-size: 11px;
                    border: none;
                    border-radius: 3px;
                }
                QPushButton:hover {
                    background-color: #777;
                }
            """)
            # Use lambda with default argument to capture current button and index
            bootloader_btn.clicked.connect(
                lambda checked, btn=bootloader_btn, idx=i: self.handle_bootloader_click(btn, idx)
            )
            
            field_layout.addWidget(label)
            field_layout.addWidget(input_field)
            field_layout.addWidget(bootloader_btn)
            
            # Bootloader indicator
            boot_led = QLabel()
            boot_led.setFixedSize(18, 18)
            boot_led.setStyleSheet("""
                QLabel {
                    background-color: #999;
                    border-radius: 9px;
                    border: 1px solid #666;
                }
            """)
            boot_led.setAlignment(Qt.AlignmentFlag.AlignCenter)
            
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


            field_layout.addWidget(boot_led)
            field_layout.addWidget(sn_led)

   # Add firmware button next to LEDs
            firmware_btn = QPushButton("For firmware")
            firmware_btn.setMaximumWidth(80)
            firmware_btn.setStyleSheet("""
                QPushButton {
                    background-color: #999;
                    color: white;
                    padding: 5px;
                    font-size: 11px;
                    border: none;
                    border-radius: 3px;
                }
                QPushButton:hover {
                    background-color: #777;
                }
            """)
            # Use lambda with default argument to capture current button and index
            firmware_btn.clicked.connect(
                lambda checked, btn=firmware_btn, idx=i: self.handle_firmware_click(btn, idx)
            )
            
            field_layout.addWidget(firmware_btn)

            scroll_layout.addLayout(field_layout)
            self.serial_inputs.append(input_field)
            self.bootloader_indicators.append(boot_led)
            self.serial_verify_indicators.append(sn_led)
            self.bootloader_buttons.append(bootloader_btn)
            self.firmware_buttons.append(firmware_btn)


            
        
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
        
        # Reset button (moved after Clear)
        reset_btn = QPushButton("Reset Power")
        reset_btn.setStyleSheet("""
            QPushButton {
                background-color: #9C27B0;
                color: white;
                padding: 10px;
                font-size: 14px;
                border: none;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #7B1FA2;
            }
        """)
        reset_btn.clicked.connect(self.send_reset_command)
        
        button_layout.addWidget(save_btn)
        button_layout.addWidget(clear_btn)
        button_layout.addWidget(reset_btn)
        main_layout.addLayout(button_layout)
        
        # Status label
        self.status_label = QLabel("No data saved")
        self.status_label.setStyleSheet("font-size: 11px; color: #666; padding: 5px;")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(self.status_label)
    

    def refresh_ports(self):
        """Refresh the list of available COM ports"""
        self.port_combo.clear()
        ports = serial.tools.list_ports.comports()
        
        if ports:
            for port in ports:
                self.port_combo.addItem(f"{port.device} - {port.description}", port.device)
        else:
            self.port_combo.addItem("No COM ports found")
    
    def connect_serial(self):
        """Connect to the selected COM port"""
        if self.port_combo.currentData():
            try:
                port = self.port_combo.currentData()
                self.serial_port = serial.Serial(
                    port=port,
                    baudrate=19200,
                    parity=serial.PARITY_NONE,
                    stopbits=serial.STOPBITS_ONE,
                    bytesize=serial.EIGHTBITS,
                    timeout=1
                )
                
                # Update UI
                self.connection_status.setText("● Connected")
                self.connection_status.setStyleSheet("color: #4CAF50; font-weight: bold;")
                self.connect_btn.setEnabled(False)
                self.disconnect_btn.setEnabled(True)
                self.port_combo.setEnabled(False)
                
                print(f"Connected to {port}")
                QMessageBox.information(self, "Connected", f"Successfully connected to {port}")
                
            except Exception as e:
                QMessageBox.critical(self, "Connection Error", f"Failed to connect to {port}\n\nError: {str(e)}")
                print(f"Connection error: {str(e)}")
        else:
            QMessageBox.warning(self, "No Port", "Please select a COM port!")
    
    def disconnect_serial(self):
        """Disconnect from the serial port"""
        if self.serial_port and self.serial_port.is_open:
            try:
                self.serial_port.close()
                
                # Update UI
                self.connection_status.setText("● Disconnected")
                self.connection_status.setStyleSheet("color: #f44336; font-weight: bold;")
                self.connect_btn.setEnabled(True)
                self.disconnect_btn.setEnabled(False)
                self.port_combo.setEnabled(True)
                
                print("Disconnected from serial port")
                
            except Exception as e:
                QMessageBox.critical(self, "Disconnection Error", f"Failed to disconnect\n\nError: {str(e)}")
                print(f"Disconnection error: {str(e)}")
    

    def handle_bootloader_click(self, button, field_index):
        """Handle bootloader button click - reset all buttons to grey, then make this one orange"""
        # Reset the last clicked button to grey if it exists
        if self.last_clicked_button is not None:
            self.last_clicked_button.setStyleSheet("""
                QPushButton {
                    background-color: #999;
                    color: white;
                    padding: 5px;
                    font-size: 11px;
                    border: none;
                    border-radius: 3px;
                }
                QPushButton:hover {
                    background-color: #777;
                }
            """)
        
        # Set this button to orange
        button.setStyleSheet("""
            QPushButton {
                background-color: #FF9800;
                color: white;
                padding: 5px;
                font-size: 11px;
                border: none;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #F57C00;
            }
        """)
        
        # Update last clicked button
        self.last_clicked_button = button
        
        # Call your original function
        self.send_serial_data_for_bootloader(field_index)


    def handle_firmware_click(self, button, field_index):
        """Handle firmware button click - reset all buttons to grey, then make this one orange"""
        # Reset the last clicked button to grey if it exists
        if self.last_clicked_button is not None:
            self.last_clicked_button.setStyleSheet("""
                QPushButton {
                    background-color: #999;
                    color: white;
                    padding: 5px;
                    font-size: 11px;
                    border: none;
                    border-radius: 3px;
                }
                QPushButton:hover {
                    background-color: #777;
                }
            """)
        
        # Set this button to orange
        button.setStyleSheet("""
            QPushButton {
                background-color: #FF9800;
                color: white;
                padding: 5px;
                font-size: 11px;
                border: none;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #F57C00;
            }
        """)
        
        # Update last clicked button
        self.last_clicked_button = button
        
        # Call your original function
        self.send_serial_data_for_firmware(field_index)

    def send_serial_data_for_bootloader(self, field_number):
        """Send serial data for a specific field number"""
        try:
            data = [0x41, 0x01, field_number, 0x0D]
            data_bytes = bytes(data)
            
            if self.serial_port and self.serial_port.is_open:
                self.serial_port.write(data_bytes)
                self.status_label.setText(f"Sent data for bootloader {field_number}")
                self.status_label.setStyleSheet("font-size: 11px; color: #FF9800; padding: 5px;")
            else:
                QMessageBox.warning(self, "Serial Error", "Serial port is not open! Please connect first.")
                
        except Exception as e:
            print(f"Error sending serial data for bootloader {field_number}: {str(e)}")
            QMessageBox.warning(
                self,
                "Send Error",
                f"Failed to send data for Serial {field_number}\n\nError: {str(e)}"
            )

    def send_serial_data_for_firmware(self, field_number):
        """Send serial data for a specific field number"""
        try:
            data = [0x41, 0x01, field_number+8, 0x0D]
            data_bytes = bytes(data)
            
            if self.serial_port and self.serial_port.is_open:
                self.serial_port.write(data_bytes)
                self.status_label.setText(f"Sent data for firmware {field_number}")
                self.status_label.setStyleSheet("font-size: 11px; color: #FF9800; padding: 5px;")
            else:
                QMessageBox.warning(self, "Serial Error", "Serial port is not open! Please connect first.")
                
        except Exception as e:
            print(f"Error sending serial data for field {field_number}: {str(e)}")
            QMessageBox.warning(
                self,
                "Send Error",
                f"Failed to send data for Serial {field_number}\n\nError: {str(e)}"
            )

    def send_reset_command(self):
        """Send reset command 0x41 0x01 0xFF 0x0D"""
        try:
            # Prepare reset command
            data = [0x41, 0x01, 0xFF, 0x0D]
            data_bytes = bytes(data)
            
            # Send via serial port
            if self.serial_port and self.serial_port.is_open:
                self.serial_port.write(data_bytes)
                print(f"Sent reset command: {' '.join([f'0x{b:02X}' for b in data_bytes])}")
                
                # Update status label
                self.status_label.setText("✓ Reset command sent")
                self.status_label.setStyleSheet("font-size: 11px; color: #9C27B0; padding: 5px;")
                
                QMessageBox.information(self, "Reset Sent", "Reset command sent successfully!")
            else:
                QMessageBox.warning(self, "Serial Error", "Serial port is not open! Please connect first.")
                
        except Exception as e:
            print(f"Error sending reset command: {str(e)}")
            QMessageBox.warning(
                self,
                "Send Error",
                f"Failed to send reset command\n\nError: {str(e)}"
            )
    
    def save_serial_numbers(self):
        # Collect all serial numbers
        serial_data = {}
        empty_count = 0
        
        for i, input_field in enumerate(self.serial_inputs, 1):
            value = input_field.text().strip()
            if value:
                serial_data[f"serial_{i}"] = value
            else:
                empty_count += 1
        
        if not serial_data:
            QMessageBox.warning(self, "No Data", "Please enter at least one serial number!")
            return
        
        # Save to memory (overwrites previous save)
        self.saved_data = {
            "data": serial_data,
            "total_entries": len(serial_data)
        }
        
        # Update status
        self.status_label.setText(f"✓ Saved {len(serial_data)} serial number(s)")
        self.status_label.setStyleSheet("font-size: 11px; color: #4CAF50; padding: 5px;")
        
        msg = f"Serial numbers saved temporarily!\n\nEntries saved: {len(serial_data)}"
        if empty_count > 0:
            msg += f"\nEmpty fields: {empty_count}"
        
        QMessageBox.information(self, "Success", msg)
        
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
                for row_idx, _ in enumerate(self.bootloader_indicators):

                    self.bootloader_indicators[row_idx].setStyleSheet("""
                                                                    background-color: #999;
                                                                    border-radius: 9px;
                                                                    border: 1px solid #666;
                                                                    """
                                                                    )        
                    self.serial_verify_indicators[row_idx].setStyleSheet("""
                                                                        background-color: #999;
                                                                        border-radius: 9px;
                                                                        border: 1px solid #666;
                                                                        """
                                                                        )

    
    def select_bootloader(self):
        """Open file dialog to select bootloader file"""
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Select Bootloader File",
            "",
            "Binary Files (*.bin);;All Files (*)"
        )
        
        if filename:
            self.bootloader_path.setText(filename)
            print(f"Bootloader selected: {filename}")
    
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
        print(f"Total Entries: {self.saved_data['total_entries']}")
        print("=" * 50)
        
        # Get file paths
        bootloader = self.bootloader_path.text()
        firmware = self.firmware_path.text()

        # Validate that paths are selected
        if not bootloader:
            QMessageBox.warning(self, "Missing File", "Please select a bootloader file!")
            return
        
        if not firmware:
            QMessageBox.warning(self, "Missing File", "Please select a firmware file!")
            return
        
        # Check if serial port is connected
        if not self.serial_port or not self.serial_port.is_open:
            QMessageBox.warning(self, "Serial Error", "Serial port is not connected! Please connect first.")
            return
        
        # Build automation queue
        self.automation_queue = []
        cycle_number = 0x01
        
        for key, serial_number in self.saved_data['data'].items():
            self.automation_queue.append({
                'key': key,
                'serial_number': serial_number,
                'cycle_number': cycle_number,
                'bootloader': bootloader,
                'firmware': firmware
            })
            cycle_number += 1
        
        # Start processing queue
        self.process_next_in_queue()
    


    def update_bootloader_status(self, row, success):
        led = self.bootloader_indicators[row]
        if success:
            led.setStyleSheet(
                """
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

    def update_serial_verify_status(self, row, success):
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
        self.bootloader_indicators[row_idx].setStyleSheet("""
            background-color: #999;
            border-radius: 9px;
            border: 1px solid #666;
        """)        
        self.serial_verify_indicators[row_idx].setStyleSheet("""
                background-color: #999;
                border-radius: 9px;
                border: 1px solid #666;
            """)

        print(f"\nProcessing {task['key']}: {task['serial_number']} (Cycle 0x{task['cycle_number']:02X})")
        
        # Update status
        self.status_label.setText(f"⏳ Processing {task['serial_number']}...")
        self.status_label.setStyleSheet("font-size: 11px; color: #2196F3; padding: 5px;")
        
        # Create and start thread
        self.current_thread = AutomationThread(
            serial_number=task['serial_number'],
            bootloader_path=task['bootloader'],
            firmware_path=task['firmware'],
            bat_file=r"D:\MULTIPROGRAMMER\flash.bat",
            driver_path=r"D:\MULTIPROGRAMMER\chromedriver-win64\chromedriver-win64\chromedriver.exe",
            chromefortestbinary_path = r"D:\MULTIPROGRAMMER\chrome-win64\chrome-win64\chrome.exe",
            cycle_number=task['cycle_number'],
            serial_port=self.serial_port,
            row_index=task['cycle_number'] - 1 
        )
        
        # Connect signals
        self.current_thread.progress.connect(self.on_automation_progress)
        self.current_thread.finished.connect(self.on_automation_finished)
        self.current_thread.bootloader_status.connect(self.update_bootloader_status)
        self.current_thread.serial_verify_status.connect(self.update_serial_verify_status)
        
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


def automate_device(serial_number, 
                    bootloader_path,
                    bat_file,
                    firmware_path, 
                    driver_path,
                    chromefortestbinary_path,
                    bootloader_callback,
                    serial_verify_callback,
                    serial_port,
                    cycle_number,
                    ):

    # bat_file = bat_file
    driver = None

    try:
        
        data_bytes_before_bootloader = bytes([0x41, 0x01, 0xFF, 0x0D])
        serial_port.write(data_bytes_before_bootloader)
        time.sleep(1)
        
        # FIRST SERIAL COMMAND - Before bootloader upload
        data_bootloader = [0x41, 0x01, cycle_number, 0x0D]
        data_bytes_bootloader = bytes(data_bootloader)
        serial_port.write(data_bytes_bootloader)
        time.sleep(1)
        
        # Upload bootloader
        command = f'"{bat_file}" "{bootloader_path}"'
        upload_bootloader = subprocess.run(
            command,
            shell=True,
            capture_output=True,
            text=True,
        )
        
        if upload_bootloader.stderr:
            print("STDERR:", upload_bootloader.stderr)
        boot_ok = "Programming Complete" in upload_bootloader.stdout and "Verification...OK" in upload_bootloader.stdout
        bootloader_callback(boot_ok)
        if not boot_ok:
            serial_verify_callback(False)
            raise Exception("Bootloader upload failed")
        
        data_bytes_before_firmware = bytes([0x41, 0x01, 0xFF, 0x0D])
        serial_port.write(data_bytes_before_firmware)
        time.sleep(2)
        
       
        # SECOND SERIAL COMMAND - Before web service/automation
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
        driver.get("http://192.168.0.100/factoryconfig")
        
        # Wait for page to be fully loaded
        WebDriverWait(driver, 5).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input[name="serialnumber"]'))
        )

        
        driver.find_element(By.CSS_SELECTOR, 'input[name="serialnumber"]').clear()
        driver.find_element(By.CSS_SELECTOR, 'input[name="serialnumber"]').send_keys(serial_number)
        driver.find_element(By.CSS_SELECTOR, 'input[type="submit"][value="Update"]').click()
        time.sleep(0.5)
        
        driver.find_element(By.XPATH, '//button[text()="Exit to bootloader"]').click()
        
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "Upload-FW"))
        )
        time.sleep(0.5)
        
        file_input = driver.find_element(By.ID, "Upload-FW")
        file_input.send_keys(firmware_path)
        driver.find_element(By.CSS_SELECTOR, "div.fws-btn.fws-btn-upload").click()
        time.sleep(0.5)
        
        # The firmware upload and device reboot takes ~15 seconds - this is hardware limitation
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
            
            print(f"Serial Number memory: {serial_number}")
            print(f"Serial Number from device: {serial_number_from_device}")
            
            sn_match = (serial_number_from_device == serial_number)
            serial_verify_callback(sn_match)
            
        except Exception as e:
            print(f"Failed to verify serial number: {e}")
            serial_verify_callback(False)
        

    except Exception as e:
        print(f"Automation error: {e}")
        serial_verify_callback(False)
    
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