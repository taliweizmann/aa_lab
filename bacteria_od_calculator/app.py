import os
import math
from flask import Flask, render_template, request
import openpyxl

app = Flask(__name__)

# Path to the Excel file in the same directory as app.py
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, "VolumeOfBacteriaBasedOnOD_clean.xlsx")
SHEET_NAME = "10^6 in 1ul"


def excel_round(val, digits=0):
    """Exact replica of Excel's ROUND() function (half away from zero)."""
    factor = 10 ** digits
    if val >= 0:
        return math.floor(val * factor + 0.5) / factor
    return math.ceil(val * factor - 0.5) / factor


def load_bacteria_data():
    """Reads the latest Excel file and extracts all 17 bacteria and their valid rows."""
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=False)
    ws = wb[SHEET_NAME]

    bacteria_list = []
    bacteria_records = {}

    for row in range(2, ws.max_row + 1):
        name = ws.cell(row, 1).value  # Col A: BACTERIA
        if not name:
            continue
        name = str(name).strip()

        if name not in bacteria_list:
            bacteria_list.append(name)
            bacteria_records[name] = []

        del_exp = ws.cell(row, 2).value   # Col B: Del:10^x
        cfu = ws.cell(row, 3).value       # Col C: cfu
        test_od = ws.cell(row, 5).value   # Col E: test OD

        # Check if numeric values exist in this row
        has_cfu = isinstance(cfu, (int, float))
        has_del = isinstance(del_exp, (int, float))
        has_od = isinstance(test_od, (int, float))

        bact_in_1ml = None
        wanted_od = None
        if has_cfu and has_del:
            # Col D formula: =C * 10 * 10^B
            bact_in_1ml = cfu * 10 * (10 ** del_exp)
            if has_od and bact_in_1ml > 0:
                # Col F formula: =(10^6 / D) * E
                wanted_od = (10 ** 6 / bact_in_1ml) * test_od

        bacteria_records[name].append({
            "row": row,
            "del_exp": del_exp,
            "cfu": cfu,
            "test_od": test_od,
            "bact_in_1ml": bact_in_1ml,
            "wanted_od": wanted_od,
            "is_ready": wanted_od is not None
        })

    return bacteria_list, bacteria_records


def calculate_protocol(wanted_od, new_od):
    """Exact implementation of columns H through Q from the Excel sheet."""
    # Col H: =(F / G) * 1000
    H = (wanted_od / new_od) * 1000.0

    # Col I & J: Tube 1
    if excel_round(H, 0) >= 20:
        I = "Take directly"
        J = None
    elif excel_round(H * 10, 0) >= 20:
        I = H * 10
        J = 1000 - I
    else:
        I = 100
        J = 1000 - I

    # Col K & L: Tube 2
    if excel_round(H * 10, 0) >= 20:
        K = "-"
        L = None
    elif excel_round(H * 100, 0) >= 20:
        K = H * 100
        L = 1000 - K
    else:
        K = 100
        L = 1000 - K

    # Col M & N: Tube 3
    if excel_round(H * 100, 0) >= 20:
        M = "-"
        N = None
    elif excel_round(H * 1000, 0) >= 20:
        M = H * 1000
        N = 1000 - M
    else:
        M = 100
        N = 1000 - M

    # Col O & P: Tube 4
    if excel_round(H * 1000, 0) >= 20:
        O = "-"
        P = None
    elif excel_round(H * 10000, 0) >= 20:
        O = H * 10000
        P = 1000 - O
    else:
        O = H * 100000
        P = 1000 - O

    # Col Q: Final Take from Last Tube (ul)
    if excel_round(H, 0) >= 20:
        Q = H
    else:
        Q = 100

    # Build step-by-step instructions based on active tubes
    steps = []
    if I == "Take directly":
        steps.append(f"Take {round(Q, 2)} µL directly from the stock culture (no serial dilution needed).")
    else:
        tubes = [
            ("Tube 1", "stock culture", I, J),
            ("Tube 2", "Tube 1", K, L),
            ("Tube 3", "Tube 2", M, N),
            ("Tube 4", "Tube 3", O, P),
        ]
        last_tube_name = "Tube 1"
        for tube_name, source_name, bact_vol, pbs_vol in tubes:
            if isinstance(bact_vol, (int, float)):
                steps.append(
                    f"{tube_name}: Add {round(pbs_vol, 2)} µL PBS + {round(bact_vol, 2)} µL from {source_name} (Total: 1000 µL)."
                )
                last_tube_name = tube_name
        steps.append(f"Final Inoculation: Take {round(Q, 2)} µL from {last_tube_name}.")

    return {
        "H": round(H, 4),
        "I": round(I, 2) if isinstance(I, (int, float)) else I,
        "J": round(J, 2) if isinstance(J, (int, float)) else "-",
        "K": round(K, 2) if isinstance(K, (int, float)) else K,
        "L": round(L, 2) if isinstance(L, (int, float)) else "-",
        "M": round(M, 2) if isinstance(M, (int, float)) else M,
        "N": round(N, 2) if isinstance(N, (int, float)) else "-",
        "O": round(O, 2) if isinstance(O, (int, float)) else O,
        "P": round(P, 2) if isinstance(P, (int, float)) else "-",
        "Q": round(Q, 2),
        "steps": steps
    }


@app.route("/", methods=["GET", "POST"])
def index():
    bacteria_list, bacteria_records = load_bacteria_data()
    selected_bacteria = ""
    new_od_input = ""
    result = None
    record_used = None
    error_msg = None

    if request.method == "POST":
        selected_bacteria = request.form.get("bacteria", "")
        new_od_input = request.form.get("new_od", "").strip()

        if selected_bacteria in bacteria_records:
            # Find the row that has CFU data filled in (prefers the most recent/bottom filled row)
            valid_rows = [r for r in bacteria_records[selected_bacteria] if r["is_ready"]]

            if not valid_rows:
                error_msg = (
                    f"No CFU data found in the Excel file for {selected_bacteria} yet. "
                    f"Please update the 'cfu' column in VolumeOfBacteriaBasedOnOD_clean.xlsx first."
                )
            else:
                record_used = valid_rows[-1]
                try:
                    new_od = float(new_od_input)
                    if new_od <= 0:
                        error_msg = "Measured OD must be greater than 0."
                    else:
                        result = calculate_protocol(record_used["wanted_od"], new_od)
                except ValueError:
                    error_msg = "Please enter a valid numeric OD value."

    return render_template(
        "index.html",
        bacteria_list=bacteria_list,
        selected_bacteria=selected_bacteria,
        new_od_input=new_od_input,
        record_used=record_used,
        result=result,
        error_msg=error_msg
    )


if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=5001)
