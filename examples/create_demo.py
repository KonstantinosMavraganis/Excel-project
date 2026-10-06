"""Create synthetic Excel files for trying the desktop application."""
from pathlib import Path

from openpyxl import Workbook


def main():
    directory = Path(__file__).resolve().parent
    datasets = {
        "demo_customers.xlsx": [
            ["Κωδικός πελάτη", "Επωνυμία", "Ενεργό συμβόλαιο"],
            ["001", "Δοκιμαστικός πελάτης Α", "Ναι"],
            ["002", "Δοκιμαστικός πελάτης Β", "Ναι"],
            ["003", "Δοκιμαστικός πελάτης Γ", "Ναι"],
            ["004", "Δοκιμαστικός πελάτης Δ", "Ναι"],
            ["005", "Ανενεργός πελάτης", "Όχι"],
        ],
        "demo_payments.xlsx": [
            ["Κωδικός πελάτη", "Ποσό"],
            ["001", 100], ["003", 150], ["001", 50], ["999", 80],
        ],
    }
    for name, rows in datasets.items():
        workbook = Workbook()
        workbook.active.title = "Δεδομένα"
        for row in rows:
            workbook.active.append(row)
        workbook.save(directory / name)
        workbook.close()
        print(name)


if __name__ == "__main__":
    main()
