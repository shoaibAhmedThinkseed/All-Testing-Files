"""
Parse BLE/SPI logic analyser export (xlsx).

The file contains different row lengths:

    enable:
        name, type, start_time, duration

    disable:
        name, type, start_time, duration

    result:
        name, type, start_time, duration, mosi, miso

Only result rows require MOSI/MISO columns.
"""

import sys
import os
import openpyxl


def parse(xlsx_path: str) -> None:

    print(f"Reading file: {xlsx_path}")

    wb = openpyxl.load_workbook(
        xlsx_path,
        read_only=True,
        data_only=True
    )

    ws = wb.active

    chunks = []

    current_mosi = []
    current_miso = []

    inside_transaction = False

    skipped_rows = 0
    result_rows = 0
    enable_count = 0
    disable_count = 0

    # ---------------------------------------------------------
    # Read every row
    # ---------------------------------------------------------

    for row_number, row in enumerate(
        ws.iter_rows(min_row=2, values_only=True),
        start=2
    ):

        # Need at least name + type
        if row is None or len(row) < 2:
            skipped_rows += 1
            continue

        name = row[0]
        typ = row[1]

        # -----------------------------------------------------
        # SPI transaction START
        # -----------------------------------------------------

        if typ == "enable":

            enable_count += 1

            inside_transaction = True

            current_mosi = []
            current_miso = []

            continue

        # -----------------------------------------------------
        # SPI transaction END
        # -----------------------------------------------------

        if typ == "disable":

            disable_count += 1

            if inside_transaction and current_mosi:

                chunks.append(
                    (
                        list(current_mosi),
                        list(current_miso)
                    )
                )

            inside_transaction = False

            current_mosi = []
            current_miso = []

            continue

        # -----------------------------------------------------
        # SPI DATA
        # -----------------------------------------------------

        if typ == "result":

            result_rows += 1

            # Result rows must have MOSI and MISO
            if len(row) < 6:

                skipped_rows += 1
                continue

            # Only collect results inside an SPI transaction
            if inside_transaction:

                mosi = row[4]
                miso = row[5]

                if mosi is None:
                    mosi = "0x??"

                if miso is None:
                    miso = "0x??"

                current_mosi.append(str(mosi))
                current_miso.append(str(miso))

    # ---------------------------------------------------------
    # If file ends while transaction is still active
    # ---------------------------------------------------------

    if inside_transaction and current_mosi:

        chunks.append(
            (
                list(current_mosi),
                list(current_miso)
            )
        )

    # ---------------------------------------------------------
    # Output filename
    # ---------------------------------------------------------

    base = os.path.splitext(xlsx_path)[0]

    out_path = base + ".txt"

    # ---------------------------------------------------------
    # Write output
    # ---------------------------------------------------------

    with open(
        out_path,
        "w",
        encoding="utf-8"
    ) as f:

        for transaction_number, (mosi_bytes, miso_bytes) in enumerate(
            chunks,
            start=1
        ):

            f.write(
                "MOSI: "
                + " ".join(mosi_bytes)
                + "\n"
            )

            f.write(
                "MISO: "
                + " ".join(miso_bytes)
                + "\n"
            )

            f.write("\n")

    # ---------------------------------------------------------
    # Print summary
    # ---------------------------------------------------------

    print()
    print("============================================")
    print("       BLE / SPI PARSING COMPLETED")
    print("============================================")
    print(f"Enable events       : {enable_count}")
    print(f"Disable events      : {disable_count}")
    print(f"Result rows         : {result_rows}")
    print(f"Transactions found  : {len(chunks)}")
    print(f"Skipped rows        : {skipped_rows}")
    print("--------------------------------------------")
    print(f"Output file:")
    print(out_path)
    print("============================================")


if __name__ == "__main__":

    if len(sys.argv) > 1:
        path = sys.argv[1]
    else:
        path = "Owner_phone_First_transaction.xlsx"

    if not os.path.exists(path):

        print()
        print("ERROR: File not found!")
        print(f"File: {path}")
        print()

        sys.exit(1)

    parse(path)