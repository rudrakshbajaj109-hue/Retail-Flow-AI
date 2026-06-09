from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os
from analysis import analyze_files

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = "backend/uploads"
REPORT_FOLDER = "backend/reports"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(REPORT_FOLDER, exist_ok=True)


@app.route("/")
def home():
    return "StockPulse AI Backend is running"


@app.route("/analyze", methods=["POST"])
def analyze():
    if "purchase_file" not in request.files:
        return jsonify({"error": "Purchase file is required"}), 400

    if "sales_file" not in request.files:
        return jsonify({"error": "Sales file is required"}), 400

    purchase_file = request.files["purchase_file"]
    sales_file = request.files["sales_file"]

    purchase_path = os.path.join(UPLOAD_FOLDER, purchase_file.filename)
    sales_path = os.path.join(UPLOAD_FOLDER, sales_file.filename)

    purchase_file.save(purchase_path)
    sales_file.save(sales_path)

    try:
        retailer_info = {
        "shop_name": request.form.get("shop_name", ""),
        "owner_name": request.form.get("owner_name", ""),
        "email": request.form.get("email", ""),
        "phone": request.form.get("phone", ""),
        "city": request.form.get("city", "")
    }

        report_path, final_df = analyze_files(purchase_path, sales_path, retailer_info)
        

        summary = {
            "total_items": int(len(final_df)),
            "high_value_items": int((final_df["VALUE CATEGORY"] == "High Value").sum()),
            "moderate_value_items": int((final_df["VALUE CATEGORY"] == "Moderate Value").sum()),
            "low_value_items": int((final_df["VALUE CATEGORY"] == "Low Value").sum()),
            "report_path": report_path
        }

        return jsonify({
            "message": "Analysis completed successfully",
            "summary": summary,
            "preview": final_df.head(10).to_dict(orient="records")
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/download-report", methods=["GET"])
def download_report():
    report_path = os.path.abspath(
        os.path.join(REPORT_FOLDER, "final_analysis_report.xlsx")
    )

    if not os.path.exists(report_path):
        return jsonify({"error": "Report not found"}), 404

    return send_file(
        report_path,
        as_attachment=True,
        download_name="StockPulse_Final_Report.xlsx"
    )


@app.route("/dashboard", methods=["GET"])
def dashboard():
    dashboard_path = os.path.abspath(
        os.path.join(REPORT_FOLDER, "stockpulse_dashboard.html")
    )

    if not os.path.exists(dashboard_path):
        return jsonify({"error": "Dashboard not found. Please analyze files first."}), 404

    return send_file(dashboard_path, mimetype="text/html")
if __name__ == "__main__":
    app.run(debug=True)