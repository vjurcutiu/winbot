from flask import Flask, send_from_directory, jsonify, request
import os
import win32com.client as win32

app = Flask(__name__, static_folder='build', static_url_path='')

@app.route('/')
def serve_index():
    return app.send_static_file('index.html')

@app.route('/<path:path>')
def serve_static(path):
    return send_from_directory(app.static_folder, path)

@app.route('/run_macro', methods=['POST'])
def run_macro():
    # Your existing macro route logic...
    if 'file' not in request.files or 'macro_name' not in request.form:
        return jsonify({"status": "error", "message": "Missing file or macro name"}), 400

    file = request.files['file']
    macro_name = request.form['macro_name']
    input_path = os.path.join("./", f"temp_{file.filename}")
    file.save(input_path)

    word = win32.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(os.path.abspath(input_path))
        word.Application.Run(macro_name)
        output_path = os.path.join("./", f"processed_{file.filename}")
        doc.SaveAs(os.path.abspath(output_path))
        doc.Close()
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500
    finally:
        word.Quit()

    return jsonify({
        "status": "success",
        "output_file": output_path,
        "log": "Macro executed successfully."
    })

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
