
from datetime import datetime
import os
import sys
import serial
import threading
from datetime import datetime
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QPushButton,
    QHBoxLayout, QComboBox, QTableWidget, QTableWidgetItem,
    QSpinBox, QLineEdit, QFileDialog, QMessageBox, QCheckBox, QDialog
)
from PySide6.QtCore import QTimer, Qt, QEvent
import pandas as pd
import serial.tools.list_ports

class ECANGui(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ECAN-U01 CAN Logger")
        self.serial_port = None
        self.running = False
        self.messages = []
        self.can1_open = False
        self.can2_open = False
        self.can1_bitrate = "500"
        self.can2_bitrate = "500"
        self.setup_ui()
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_table)
        self.timer.start(500)

    def default_filename(self, prefix):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return f"{prefix}_{timestamp}.xlsx"

    def eventFilter(self, source, event):
        if source == self.port_box and event.type() == QEvent.MouseButtonPress:
            self.refresh_serial_ports()
        return super().eventFilter(source, event)

    def refresh_serial_ports(self):
        self.port_box.clear()
        ports = serial.tools.list_ports.comports()
        for port in ports:
            self.port_box.addItem(port.device)

    def update_data_field(self):
        length = self.length_spin.value()
        raw = self.data_input.text().replace(" ", "").upper()
        raw = ''.join(c for c in raw if c in "0123456789ABCDEF")
        raw = raw[:length * 2]
        grouped = [raw[i:i+2] for i in range(0, len(raw), 2)]
        formatted = " ".join(grouped)
        self.data_input.blockSignals(True)
        self.data_input.setText(formatted)
        self.data_input.blockSignals(False)
        self.data_input.setCursorPosition(len(formatted))

    def update_frame_id_field(self):
        raw = self.frame_id_input.text().upper()
        filtered = ''.join(c for c in raw if c in "0123456789ABCDEF")

        frame_type = self.frame_type_box.currentText()
        try:
            if filtered:
                val = int(filtered, 16)
                if frame_type == "Standard":
                    if val > 0x7FF:
                        val = 0x7FF
                    filtered = f"{val:03X}"
                elif frame_type == "Extended":
                    if val > 0x1FFFFFFF:
                        val = 0x1FFFFFFF
                    filtered = f"{val:08X}"
        except ValueError:
            filtered = ""

        self.frame_id_input.setText(filtered)
        self.frame_id_input.setCursorPosition(len(filtered))

    def setup_ui(self):
        layout = QVBoxLayout()

        # Serial config
        serial_layout = QHBoxLayout()
        self.port_box = QComboBox()
        self.refresh_serial_ports()
        self.baudrate_box = QComboBox()
        self.baudrate_box.addItems(["9600", "115200"])
        self.open_btn = QPushButton("Open")
        self.open_btn.clicked.connect(self.toggle_connection)
        self.port_box.installEventFilter(self)
        serial_layout.addWidget(QLabel("Port:"))
        serial_layout.addWidget(self.port_box)
        serial_layout.addWidget(QLabel("Baudrate:"))
        serial_layout.addWidget(self.baudrate_box)
        serial_layout.addWidget(self.open_btn)
        layout.addLayout(serial_layout)

        # CAN channel controls
        can_control_layout = QHBoxLayout()
        self.can1_btn = QPushButton("Open CAN1")
        self.can1_btn.clicked.connect(lambda: self.toggle_can(1))
        self.can1_speed = QComboBox()
        self.can1_speed.addItems(["500", "250", "125"])
        self.can2_btn = QPushButton("Open CAN2")
        self.can2_btn.clicked.connect(lambda: self.toggle_can(2))
        self.can2_speed = QComboBox()
        self.can2_speed.addItems(["500", "250", "125"])
        can_control_layout.addWidget(QLabel("CAN1 Baudrate (Kbps):"))
        can_control_layout.addWidget(self.can1_speed)
        can_control_layout.addWidget(self.can1_btn)
        can_control_layout.addWidget(QLabel("CAN2 Baudrate (Kbps):"))
        can_control_layout.addWidget(self.can2_speed)
        can_control_layout.addWidget(self.can2_btn)
        layout.addLayout(can_control_layout)

        # Channel selection for display
        channel_display_layout = QHBoxLayout()
        self.show_can1 = QCheckBox("Show CAN1 Messages")
        self.show_can1.setChecked(True)
        self.show_can2 = QCheckBox("Show CAN2 Messages")
        self.show_can2.setChecked(True)
        channel_display_layout.addWidget(self.show_can1)
        channel_display_layout.addWidget(self.show_can2)
        layout.addLayout(channel_display_layout)

        # CAN send
        send_layout = QHBoxLayout()
        self.channel_box = QComboBox()
        self.channel_box.addItems(["1", "2"])
        self.frame_id_input = QLineEdit("")
        self.frame_id_input.textEdited.connect(self.update_frame_id_field)
        self.length_spin = QSpinBox()
        self.length_spin.setRange(0, 8)
        self.length_spin.setValue(8)
        self.length_spin.valueChanged.connect(self.update_data_field)
        self.data_input = QLineEdit("")
        self.data_input.textEdited.connect(self.update_data_field)
        self.send_btn = QPushButton("Send")
        self.send_btn.clicked.connect(self.send_message)
        self.frame_type_box = QComboBox()
        self.frame_type_box.addItems(["Standard", "Extended"])
        send_layout.addWidget(QLabel("Frame Type:"))
        send_layout.addWidget(self.frame_type_box)
        send_layout.addWidget(QLabel("Channel:"))
        send_layout.addWidget(self.channel_box)
        send_layout.addWidget(QLabel("FrameID:"))
        send_layout.addWidget(self.frame_id_input)
        send_layout.addWidget(QLabel("Lenght:"))
        send_layout.addWidget(self.length_spin)
        send_layout.addWidget(QLabel("Data:"))
        send_layout.addWidget(self.data_input)
        send_layout.addWidget(self.send_btn)
        layout.addLayout(send_layout)

        # Table
        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(["Direction", "Channel", "FrameID", "Lenght", "Data", "Timestamp"])
        layout.addWidget(self.table)
        self.table.setColumnWidth(4, 250) 
        self.table.setColumnWidth(5, 200) 
        
        # Group and count button
        self.group_count_btn = QPushButton("Group and Count Messages")
        self.group_count_btn.clicked.connect(self.group_and_count_messages)
        layout.addWidget(self.group_count_btn)

        # Save button
        self.save_btn = QPushButton("Save to XLSX")
        self.save_btn.clicked.connect(self.save_log)
        layout.addWidget(self.save_btn)

        self.setLayout(layout)

    def toggle_connection(self):
        if self.serial_port:
            self.running = False
            self.serial_port.close()
            self.serial_port = None
            self.open_btn.setText("Open")
            self.can1_open = False
            self.can2_open = False
            self.can1_btn.setText("Open CAN1")
            self.can2_btn.setText("Open CAN2")
        else:
            try:
                self.serial_port = serial.Serial(
                    port=self.port_box.currentText(),
                    baudrate=int(self.baudrate_box.currentText()),
                    timeout=0.1
                )
                self.serial_port.reset_input_buffer()
                self.serial_port.write(bytes.fromhex("f0 01 0d 0a"))
                response = self.serial_port.read(20)
                if response.startswith(b'\xff\x01\x01\x02'):
                    self.running = True
                    threading.Thread(target=self.read_loop, daemon=True).start()
                    self.open_btn.setText("Close")
                    QMessageBox.information(self, "Handshake OK", "Device responded to handshake.")
                else:
                    self.serial_port.close()
                    self.serial_port = None
                    QMessageBox.warning(self, "Handshake Failed", "No valid response from device.")
            except Exception as e:
                self.serial_port = None
                self.open_btn.setText("Open")
                QMessageBox.critical(self, "Serial Error", f"Failed to open serial port: {e}")

    def close_handshake(self):
        if self.serial_port:
            self.serial_port.write(bytes.fromhex("f0 0f 0d 0a"))

    def toggle_can(self, channel):
        if not self.serial_port:
            return

        def build_command(ch, bitrate):
            if bitrate == "500":
                return bytes.fromhex(f"f0 02 0{ch} 08 00 0c 03 0d 0a")
            elif bitrate == "250":
                return bytes.fromhex(f"f0 02 0{ch} 12 00 0b 02 0d 0a")
            elif bitrate == "125":
                return bytes.fromhex(f"f0 02 0{ch} 24 00 0b 02 0d 0a")  
            else:
                return None

        def build_close(ch):
            return bytes.fromhex(f"f0 02 0{ch} 00 00 00 00 0d 0a")

        if channel == 1:
            if self.can1_open:
                self.serial_port.write(build_close(1))
                self.can1_open = False
                self.can1_btn.setText("Open CAN1")
            else:
                self.can1_bitrate = self.can1_speed.currentText()
                self.serial_port.write(build_command(1, self.can1_bitrate))
                self.can1_open = True
                self.can1_btn.setText("Close CAN1")
        elif channel == 2:
            if self.can2_open:
                self.serial_port.write(build_close(2))
                self.can2_open = False
                self.can2_btn.setText("Open CAN2")
            else:
                self.can2_bitrate = self.can2_speed.currentText()
                self.serial_port.write(build_command(2, self.can2_bitrate))
                self.can2_open = True
                self.can2_btn.setText("Close CAN2")

    def send_message(self):
        try:
            ch = int(self.channel_box.currentText())
            if (ch == 1 and not self.can1_open) or (ch == 2 and not self.can2_open):
                QMessageBox.warning(self, "CAN Channel Closed", f"CAN{ch} is closed. Cannot send.")
                return

            frame_id_text = self.frame_id_input.text().strip()
            if not frame_id_text:
                QMessageBox.warning(self, "Missing Frame ID", "Frame ID is required.")
                return

            frame_type = self.frame_type_box.currentText()
            if frame_type == "Standard":
                frame_id_val = int(frame_id_text, 16)
                if frame_id_val > 0x7FF:
                    QMessageBox.warning(self, "Invalid Frame ID", "Standard Frame ID must be ≤ 0x7FF.")
                    return
                frame_id_bytes = frame_id_val.to_bytes(2, byteorder='little') + b'\x00\x00'  # <-- fix aqui
            elif frame_type == "Extended":
                frame_id_val = int(frame_id_text, 16)
                if frame_id_val > 0x1FFFFFFF:
                    QMessageBox.warning(self, "Invalid Frame ID", "Extended Frame ID must be ≤ 0x1FFFFFFF.")
                    return
                frame_id_bytes = frame_id_val.to_bytes(4, byteorder='little')

            length = self.length_spin.value()
            if frame_type == "Extended":
                length |= 0x80 

            # Trata os dados digitados e remove espaços
            hex_str = ''.join(c for c in self.data_input.text().strip().upper() if c in "0123456789ABCDEF")

            # Se número ímpar de caracteres, completa com 0 à direita (ex: '4' → '40')
            if len(hex_str) % 2 == 1:
                hex_str += '0'

            # Converte em lista de inteiros (bytes)
            data_bytes = [int(hex_str[i:i+2], 16) for i in range(0, min(len(hex_str), length * 2), 2)]

            # Preenche com 0x00 até alcançar o length desejado
            data_padded = data_bytes + [0x00] * (8 - len(data_bytes))

            # Monta a mensagem
            payload = [0xf0, 0x05, ch, length] + list(frame_id_bytes) + data_padded + [0x78, 0x46, 0x23, 0x01]

            # Envia
            self.serial_port.write(bytes(payload))

            # Log dos dados realmente enviados (com padding)
            data_str_sent = " ".join(f"{b:02X}" for b in data_padded)
            self.log_message("TX", ch, f"0x{frame_id_val:X}", length, data_str_sent)

        except Exception as e:
            QMessageBox.critical(self, "Send Error", f"An error occurred while sending: {e}")

    def read_loop(self):
        while self.running:
            try:
                if self.serial_port.in_waiting:
                    data = self.serial_port.read(20)
                    if len(data) >= 13 and data[0] == 0xff:
                        frame_type = data[1]

                        # Aceita apenas ff 04 ou ff 05 com sufixo correto
                        if frame_type == 0x04:
                            pass
                        elif frame_type == 0x05:
                            if not data.endswith(b'\x78\x46\x23'):
                                continue  # Ignora ff 05 inválido
                        else:
                            continue  # Ignora outros tipos
                        
                        ch = data[2]
                        if (ch == 1 and not self.can1_open) or (ch == 2 and not self.can2_open):
                            continue
                        length = data[3]
                        if length >= 8:
                            frame_id = int.from_bytes(data[4:8], byteorder='little')
                        else:
                            frame_id = int.from_bytes(data[4:6], byteorder='little')
                        payload = " ".join(f"{data[i]:02X}" for i in range(8, min(8 + 8, len(data))))
                        self.log_message("RX", ch, f"0x{frame_id:X}", length, payload)
            except Exception as e:
                QMessageBox.critical(self, "Read Error", f"An error occurred while reading: {e}")

    def log_message(self, direction, channel, frame_id, length, data):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.messages.append({
            "Direction": direction,
            "Channel": channel,
            "FrameID": frame_id,
            "Length": length,
            "Data": data,
            "Timestamp": timestamp
        })

    def update_table(self):
        show_can1 = self.show_can1.isChecked()
        show_can2 = self.show_can2.isChecked()
        filtered_messages = [
            msg for msg in self.messages
            if (show_can1 and msg["Channel"] == 1) or (show_can2 and msg["Channel"] == 2)
        ]
        self.table.setRowCount(len(filtered_messages))
        for i, msg in enumerate(filtered_messages):
            self.table.setItem(i, 0, QTableWidgetItem(msg["Direction"]))
            self.table.setItem(i, 1, QTableWidgetItem(str(msg["Channel"])))
            self.table.setItem(i, 2, QTableWidgetItem(msg["FrameID"]))
            self.table.setItem(i, 3, QTableWidgetItem(str(msg["Length"])))
            self.table.setItem(i, 4, QTableWidgetItem(msg["Data"]))
            self.table.setItem(i, 5, QTableWidgetItem(msg["Timestamp"]))

    def group_and_count_messages(self):
        if not self.messages:
            QMessageBox.information(self, "No Data", "There are no messages to group and count.")
            return

        df = pd.DataFrame(self.messages)
        required_columns = ["Channel", "FrameID", "Data"]
        if not all(column in df.columns for column in required_columns):
            QMessageBox.warning(self, "Column Mismatch", "The DataFrame does not contain the expected columns.")
            return

        grouped = df.groupby(required_columns).size().reset_index(name='Count')
        self.grouped_df = grouped
        self.show_grouped_messages(grouped)

    def show_grouped_messages(self, grouped_df):
        from PySide6.QtWidgets import QDialog

        self.grouped_dialog = QDialog(self)
        self.grouped_dialog.setWindowTitle("Grouped Messages")
        self.grouped_dialog.setModal(True)

        dialog_layout = QVBoxLayout()

        # Tabela com dados agrupados
        table = QTableWidget(grouped_df.shape[0], grouped_df.shape[1])
        table.setHorizontalHeaderLabels(grouped_df.columns)
        for i, row in grouped_df.iterrows():
            for j, val in enumerate(row):
                table.setItem(i, j, QTableWidgetItem(str(val)))
        dialog_layout.addWidget(table)

        # Botão de salvar
        save_button = QPushButton("Save to XLSX")
        save_button.clicked.connect(lambda: self.save_grouped_dataframe(grouped_df))
        dialog_layout.addWidget(save_button)

        self.grouped_dialog.setLayout(dialog_layout)
        self.grouped_dialog.resize(600, 400)
        self.grouped_dialog.exec() 

    def save_grouped_dataframe(self, df):
        default_name = self.default_filename("grouped_log")
        path, _ = QFileDialog.getSaveFileName(self, "Save Grouped XLSX", default_name, "Excel Files (*.xlsx)")
        if path:
            try:
                df.to_excel(path, index=False)
                QMessageBox.information(self, "Success", f"Grouped messages saved to:\n{path}")
            except Exception as e:
                QMessageBox.critical(self, "Save Error", f"Failed to save file:\n{e}")

    def save_log(self):
        if not self.messages:
            return
        default_name = self.default_filename("can_log")
        path, _ = QFileDialog.getSaveFileName(self, "Save XLSX", default_name, "Excel Files (*.xlsx)")

        if path:
            df = pd.DataFrame(self.messages)
            df.to_excel(path, index=False)
   
    def closeEvent(self, event):
        self.close_handshake()
        self.running = False
        if self.serial_port and self.serial_port.is_open:
            self.serial_port.close()
        event.accept() 

if __name__ == "__main__":
    app = QApplication(sys.argv)
    gui = ECANGui()
    gui.resize(1000, 600)
    gui.show()
    sys.exit(app.exec())
