from flask import Flask, request, send_file, jsonify
import pandas as pd
import os
import traceback
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = "uploads"
DOWNLOAD_FOLDER = "downloads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

def clean_dataframe(df):
    """Removes unnamed columns and empty rows."""
    df = df.dropna(how="all")  # Remove completely empty rows
    df = df.loc[:, ~df.columns.str.contains("^Unnamed")]  # Remove unnamed columns
    return df

@app.route("/upload", methods=["POST"])
def upload_files():
    try:
        if "excel1" not in request.files or "excel2" not in request.files:
            return jsonify({"error": "Both Excel1 and Excel2 files are required"}), 400
        
        excel1 = request.files["excel1"]
        excel2 = request.files["excel2"]
        
        excel1_path = os.path.join(UPLOAD_FOLDER, "excel1.xlsx")
        
        # Get original Excel2 filename (without extension)
        excel2_filename = os.path.splitext(excel2.filename)[0]  # Extract filename without extension
        excel2_path = os.path.join(UPLOAD_FOLDER, excel2.filename)
        
        excel1.save(excel1_path)
        excel2.save(excel2_path)

        print(f"✅ Excel1 saved at: {excel1_path}")
        print(f"✅ Excel2 saved at: {excel2_path}")

        # Load and clean Excel files
        try:
            df1 = pd.read_excel(excel1_path, engine="openpyxl")
            df2 = pd.read_excel(excel2_path, engine="openpyxl")

            df1 = clean_dataframe(df1)
            df2 = clean_dataframe(df2)

            print("✅ Excel files loaded and cleaned successfully!")
        except Exception as e:
            print("❌ Error reading Excel files:", e)
            return jsonify({"error": f"Error reading Excel files: {str(e)}"}), 500
        
        # Ensure columns are in lowercase for case-insensitive matching
        df1.columns = df1.columns.str.lower()
        df2.columns = df2.columns.str.lower()
        
        required_columns = ["item", "description", "materials", "c/kg", "rate"]
        for col in ["c/kg", "rate"]:
            if col not in df1.columns or col not in df2.columns:
                return jsonify({"error": f"Required column '{col}' missing in one of the files"}), 400
        
        merge_columns = ["item", "description", "materials"]
        df_merged = df2.merge(df1, on=merge_columns, how="left", suffixes=("", "_filled"))
        
        # Update logic
        df_merged["c/kg"] = df_merged["c/kg"].combine_first(df_merged.get("c/kg_filled"))
        df_merged.loc[df_merged["item"].str.lower() == "pipe", "rate"] = ""
        df_merged.loc[df_merged["item"].str.lower() != "pipe", "rate"] = df_merged["rate"].combine_first(df_merged.get("rate_filled"))
        
        # Drop extra columns
        df_merged = df_merged[df2.columns]

        # Save processed file with original name + "_filtered.xlsx"
        processed_file_name = f"{excel2_filename}_filtered.xlsx"
        processed_file_path = os.path.join(DOWNLOAD_FOLDER, processed_file_name)
        df_merged.to_excel(processed_file_path, index=False)

        print(f"✅ Processed file saved as: {processed_file_name}")

        return jsonify({"message": "Processing completed", "download_url": f"/download/{processed_file_name}"})
    except Exception as e:
        print("❌ ERROR in /upload:", traceback.format_exc())
        return jsonify({"error": f"Internal server error: {str(e)}"}), 500

@app.route("/download/<filename>", methods=["GET"])
def download_file(filename):
    try:
        processed_file_path = os.path.join(DOWNLOAD_FOLDER, filename)
        
        print(f"📂 Checking if processed file exists: {processed_file_path}")
        if not os.path.exists(processed_file_path):
            print("❌ Processed file not found!")
            return jsonify({"error": "Processed file not found"}), 404
        
        print("✅ Processed file found! Downloading...")
        return send_file(processed_file_path, as_attachment=True)
    except Exception as e:
        print("❌ ERROR in /download:", traceback.format_exc())
        return jsonify({"error": f"Error downloading file: {str(e)}"}), 500

if __name__ == "__main__":
    print("🚀 Flask server is running on http://localhost:8000")
    app.run(debug=True, port=8000)
