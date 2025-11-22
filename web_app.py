from flask import Flask, render_template, request, send_file, redirect, url_for
import os
import tempfile
from X import calculate_reorder_quantities

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "GET":
        return render_template("index.html")

    # POST: files uploaded
    # Create a temp folder for this run
    tmpdir = tempfile.mkdtemp()
    sales_path = os.path.join(tmpdir, "sales.xlsx")
    inventory_path = os.path.join(tmpdir, "inventory.xlsx")
    ignore_path = os.path.join(tmpdir, "ignore.xlsx")
    irc_path = os.path.join(tmpdir, "IRC.xlsx")

    # Save required file: sales
    sales_file = request.files.get("sales_file")
    if not sales_file or sales_file.filename == "":
        return "Sales file is required", 400
    sales_file.save(sales_path)

    # Optional files
    inventory_file = request.files.get("inventory_file")
    if inventory_file and inventory_file.filename != "":
        inventory_file.save(inventory_path)
    else:
        inventory_path = "inventory.xlsx"  # will trigger your file-not-found logic if used

    ignore_file = request.files.get("ignore_file")
    if ignore_file and ignore_file.filename != "":
        ignore_file.save(ignore_path)
    else:
        ignore_path = "ignore.xlsx"

    irc_file = request.files.get("irc_file")
    if irc_file and irc_file.filename != "":
        irc_file.save(irc_path)
    else:
        irc_path = "IRC.xlsx"

    # Run your logic – auto_export=True skips the prompt and always generates the report
    cwd = os.getcwd()
    try:
        # run function with the paths we just saved
        calculate_reorder_quantities(
            sales_file=sales_path,
            inventory_file=inventory_path,
            ignore_file=ignore_path,
            irc_file=irc_path,
            auto_export=True,
        )

        # Your script saves "reorder_report.xlsx" in current working dir.
        # To keep everything inside tmpdir, we can look for it there OR just move it.
        report_path = os.path.join(cwd, "reorder_report.xlsx")
        if not os.path.exists(report_path):
            # maybe it saved in tmpdir instead, fallback
            report_path = os.path.join(tmpdir, "reorder_report.xlsx")

        if not os.path.exists(report_path):
            return "Could not find generated report.", 500

        # Send file to browser
        return send_file(
            report_path,
            as_attachment=True,
            download_name="reorder_report.xlsx"
        )

    finally:
        # you can choose to clean tmpdir after send_file if you want with shutil.rmtree
        pass


if __name__ == "__main__":
    # debug=True for development
    app.run(host="0.0.0.0", port=5000, debug=True)
