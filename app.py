from flask import Flask, jsonify, request, render_template_string, send_file, url_for
from PIL import Image
import os

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# HTML テンプレート
template = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel to HTML Table Converter and Image Compressor</title>
    <script>
        function updatePreview() {
            const code = document.getElementById('html_code').value;
            document.getElementById('preview').innerHTML = code;
        }

        function uploadImage(event) {
            event.preventDefault();
            const formData = new FormData();
            const image = document.querySelector('input[name="image"]').files[0];
            formData.append("image", image);
            formData.append("action", "compress");

            fetch("/", {
                method: "POST",
                body: formData
            })
            .then(response => response.json())
            .then(data => {
                if (data.compressed_image_url) {
                    const link = document.getElementById('compressed_image_link');
                    link.href = data.compressed_image_url;
                    document.getElementById('compressed_image_section').style.display = 'block';
                }
            })
            .catch(error => console.error('Error:', error));
        }
    </script>
</head>
<body>
    <h1>Excel to HTML Table Converter</h1>
    <form method="POST" enctype="multipart/form-data">
        <textarea name="excel_data" rows="10" cols="50" placeholder="Paste your Excel data here"></textarea><br>
        <button type="submit" name="action" value="convert">Convert</button>
    </form>
    {% if table %}
    <h2>Generated HTML Table</h2>
    <form method="POST">
        <textarea id="html_code" name="html_code" rows="10" cols="50" oninput="updatePreview()">{{ table }}</textarea><br>
        <button type="submit" name="action" value="preview">Update Preview</button>
    </form>
    <h2>Preview</h2>
    <div id="preview">{{ table|safe }}</div>
    {% endif %}

    <h1>Image Compressor</h1>
    <form onsubmit="uploadImage(event)">
        <input type="file" name="image" accept="image/*"><br><br>
        <button type="submit">Upload and Compress</button>
    </form>
    <div id="compressed_image_section" style="display:none;">
        <h2>Compressed Image</h2>
        <a id="compressed_image_link" href="#" download>Download Compressed Image</a>
    </div>
</body>
</html>
"""

@app.route("/", methods=["GET", "POST"])
def index():
    table = None
    compressed_image = None
    compressed_image_url = None

    if request.method == "POST":
        action = request.form.get("action")
        if action == "convert":
            excel_data = request.form.get("excel_data")
            if excel_data:
                table = convert_to_html_table(excel_data)
        elif action == "preview":
            table = request.form.get("html_code")
        elif action == "compress" and 'image' in request.files:
            image = request.files["image"]
            if image:
                compressed_image = compress_image(image)
                compressed_image_url = url_for('download_file', filename=compressed_image)

        if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
            return jsonify({'compressed_image': compressed_image, 'compressed_image_url': compressed_image_url})

    return render_template_string(template, table=table, compressed_image=compressed_image)

def convert_to_html_table(data):
    rows = data.strip().split("\n")
    table = "<table border='1' style='border-collapse: collapse;'>\n"
    for row in rows:
        table += "  <tr>\n"
        cells = row.split("\t")
        for cell in cells:
            table += f"    <td style='padding: 5px; border: 1px solid black;'>{cell}</td>\n"
        table += "  </tr>\n"
    table += "</table>\n"
    return table

def compress_image(image):
    filename = os.path.join(app.config['UPLOAD_FOLDER'], image.filename)
    image.save(filename)
    compressed_filename = os.path.join(app.config['UPLOAD_FOLDER'], "compressed_" + image.filename)
    
    with Image.open(filename) as img:
        img.save(compressed_filename, optimize=True, quality=85)  # 圧縮を実行

    return "compressed_" + image.filename

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(os.path.join(app.config['UPLOAD_FOLDER'], filename), as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
