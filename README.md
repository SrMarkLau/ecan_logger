## 🛠 ECAN-U01 CAN Logger (PySide6)

This is a graphical application built with **Python** and **PySide6 (Qt for Python)** to communicate with and monitor CAN messages using the **ECAN-U01** device.

### ✨ Features

- Automatic detection of available serial ports
- Open/close control for **CAN1** and **CAN2** channels
- Supports 125, 250, and 500 kbps bitrates
- Send **Standard** and **Extended** CAN frames with automatic zero-padding
- Real-time display of incoming messages
- Group and count repeated messages
- Export full logs and grouped data to **Excel (.xlsx)**
- Clean, user-friendly interface with **copy-paste** support from the table

### 🚀 Requirements

- Python 3.9+
- PySide6
- pandas
- pyserial

### 📦 Installation

```bash
python can_logger.py
```
