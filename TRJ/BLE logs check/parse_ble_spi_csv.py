"""
Parse a BLE/SPI logic analyser CSV export and write each SPI transaction
as a single horizontal line in a text file.

Usage:
    python parse_ble_spi.py "my_capture.csv"

Output:
    A .txt file with the same base name as the input, e.g.:
    TRJ_BLE_SPI_OP success.csv -> TRJ_BLE_SPI_OP success.txt

Expected CSV columns:
    Column 1: name
    Column 2: type
    Column 6: mosi
    Column 7: miso

Each SPI transaction (chip-select pulse) is written as:
    MOSI: 0x4A 0x00 0x60 ...
    MISO: 0x4A 0x00 0x60 ...

A blank line separates transactions.
"""

import sys
import os
import csv


def parse(csv_path: str) -> None:
    chunks: list[tuple[list[str], list[str]]] = []
    current_mosi: list[str] = []
    current_miso: list[str] = []
    inside_transaction = False

    with open(csv_path, "r", encoding="utf-8-sig", newline="") as csv_file:
        reader = csv.reader(csv_file)

        # Skip the header row, matching the original XLSX parser.
        next(reader, None)

        for row in reader:
            # Ignore incomplete rows.
            if len(row) < 7:
                continue

            name, typ = row[0], row[1]
            mosi, miso = row[5], row[6]

            if typ == "enable":
                inside_transaction = True
                current_mosi = []
                current_miso = []

            elif typ == "disable":
                if inside_transaction and current_mosi:
                    chunks.append((list(current_mosi), list(current_miso)))
                inside_transaction = False

            elif typ == "result" and inside_transaction:
                current_mosi.append(mosi if mosi else "0x??")
                current_miso.append(miso if miso else "0x??")

    # Flush any open transaction at end of file.
    if inside_transaction and current_mosi:
        chunks.append((current_mosi, current_miso))

    base = os.path.splitext(csv_path)[0]
    out_path = base + ".txt"

    with open(out_path, "w", encoding="utf-8") as f:
        for mosi_bytes, miso_bytes in chunks:
            f.write("MOSI: " + " ".join(mosi_bytes) + "\n")
            f.write("MISO: " + " ".join(miso_bytes) + "\n")
            f.write("\n")

    print(f"Written {len(chunks)} transactions to: {out_path}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python parse_ble_spi.py <input.csv>")
        sys.exit(1)

    path = sys.argv[1]
    parse(path)
