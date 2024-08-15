import pandas as pd
from PyQt5 import QtWidgets
import pyautogui
import pyperclip
import time


class MainWindow(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Automated Message Sender")

        layout = QtWidgets.QVBoxLayout()

        self.upload_button = QtWidgets.QPushButton("Upload Excel File")
        self.upload_button.clicked.connect(self.upload_excel)

        layout.addWidget(self.upload_button)
        self.setLayout(layout)

    def upload_excel(self):
        file_dialog = QtWidgets.QFileDialog.getOpenFileName(
            self, "Open Excel File", "", "Excel Files (*.xls *.xlsx)"
        )
        file_path = file_dialog[0]
        if file_path:
            df = pd.read_excel(file_path)
            self.process_data(df)

    def process_data(self, df):
        # Loop through each row in the DataFrame
        for index, row in df.iterrows():
            # Move to cursor position (x=363, y=60) and click
            pyautogui.moveTo(363, 60, duration=0.5)
            pyautogui.click()
            time.sleep(1)

            # Copy the text in "wa" column to clipboard
            wa_text = str(row["wa"])
            pyperclip.copy(wa_text)

            # Paste the text from clipboard
            pyautogui.hotkey("ctrl", "v")  # Paste using keyboard shortcut
            time.sleep(2)
            # Press Enter
            pyautogui.press("enter")
            time.sleep(5)

            # Move to cursor position (x=524, y=1001) and click
            pyautogui.moveTo(524, 1001, duration=0.5)
            pyautogui.click()

            # Copy the text in "pesan" column to clipboard
            pesan_text = str(row["pesan"])
            pyperclip.copy(pesan_text)
            time.sleep(2)
            # Paste the text from clipboard
            pyautogui.hotkey("ctrl", "v")  # Paste using keyboard shortcut
            time.sleep(2)
            # Press Enter
            pyautogui.press("enter")

            # Wait for 2 seconds
            time.sleep(2)


if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
