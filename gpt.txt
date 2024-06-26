from flask import Flask, request, render_template_string, send_file, jsonify, url_for
from PIL import Image
import os

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# HTMLテンプレート
template = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel to HTML Table Converter and Image Compressor</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/ace/1.4.12/ace.js" crossorigin="anonymous"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/ace/1.4.12/ext-searchbox.js" crossorigin="anonymous"></script>
    <style>
        body {
            font-family: Arial, sans-serif;
            display: flex;
            flex-direction: column;
            align-items: center;
        }

        h1 {
            color: #333;
        }

        .editor-container {
            display: flex;
            width: 100%;
            justify-content: space-around;
        }

        .editor-section, .preview-section {
            width: 45%;
        }

        #editor, #css-editor {
            height: 200px;
            width: 100%;
            border: 1px solid #ddd;
            border-radius: 4px;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
        }

        .ace_editor {
            font-size: 14px;
        }

        .ace-monokai {
            background-color: #272822;
        }

        button {
            background-color: #4CAF50;
            color: white;
            padding: 10px 20px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 16px;
            margin-top: 10px;
        }

        button:hover {
            background-color: #45a049;
        }

        .file-upload {
            display: none;
        }

        .file-upload-label {
            background-color: #4CAF50;
            color: white;
            padding: 10px 20px;
            border-radius: 4px;
            cursor: pointer;
            font-size: 16px;
            display: inline-block;
        }

        .file-upload-label:hover {
            background-color: #45a049;
        }

        .file-info {
            margin-top: 10px;
            font-size: 14px;
            color: #555;
        }

        .download-link {
            display: inline-block;
            margin-top: 10px;
            padding: 10px 20px;
            background-color: #4CAF50;
            color: white;
            text-decoration: none;
            border-radius: 4px;
            cursor: pointer;
        }

        .download-link:hover {
            background-color: #45a049;
        }

        .size-info {
            margin-top: 10px;
            font-size: 14px;
            color: #555;
        }

        #preview {
            width: 100%;
            height: 200px;
            border: 1px solid #ddd;
            border-radius: 4px;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
            padding: 10px;
            background-color: #fff;
            overflow: auto;
        }
    </style>
    <script>
        function updatePreview() {
            const htmlCode = ace.edit("editor").getValue();
            const cssCode = ace.edit("css-editor").getValue();
            const previewFrame = document.getElementById('preview');
            const previewDocument = previewFrame.contentDocument || previewFrame.contentWindow.document;
            previewDocument.open();
            previewDocument.write(`<style>${cssCode}</style>${htmlCode}`);
            previewDocument.close();
        }

        function uploadImage(event) {
            event.preventDefault();
            const formData = new FormData();
            const imageInput = document.querySelector('input[name="image"]');
            const image = imageInput.files[0];
            formData.append("image", image);
            formData.append("action", "compress");

            fetch("/", {
                method: "POST",
                body: formData,
                headers: {
                    'Accept': 'application/json'
                }
            })
            .then(response => response.json())
            .then(data => {
                if (data.compressed_image_url) {
                    const downloadLink = document.getElementById('compressed_image_link');
                    downloadLink.href = data.compressed_image_url;
                    document.getElementById('compressed_image_section').style.display = 'block';
                    
                    const originalSize = image.size;
                    const compressedSize = data.compressed_image_size;
                    const sizeInfo = document.getElementById('size_info');
                    sizeInfo.innerHTML = `Original Size: ${(originalSize / 1024).toFixed(2)} KB<br>Compressed Size: ${(compressedSize / 1024).toFixed(2)} KB<br>Reduction: ${((1 - compressedSize / originalSize) * 100).toFixed(2)}%`;
                }
            })
            .catch(error => console.error('Error:', error));
        }

        function showFileName() {
            const imageInput = document.querySelector('input[name="image"]');
            const fileName = imageInput.files[0].name;
            const fileInfo = document.getElementById('file_info');
            fileInfo.textContent = `Selected file: ${fileName}`;
        }

        document.addEventListener("DOMContentLoaded", function() {
            const htmlEditor = ace.edit("editor");
            htmlEditor.setTheme("ace/theme/monokai");
            htmlEditor.session.setMode("ace/mode/html");

            const cssEditor = ace.edit("css-editor");
            cssEditor.setTheme("ace/theme/monokai");
            cssEditor.session.setMode("ace/mode/css");

            const textarea = document.getElementById('html_code');
            htmlEditor.getSession().on('change', function() {
                textarea.value = htmlEditor.getValue();
            });

            htmlEditor.commands.addCommand({
                name: "replace",
                bindKey: { win: "Ctrl-R", mac: "Command-Option-F" },
                exec: function(htmlEditor) {
                    ace.require("ace/ext/searchbox").Search(htmlEditor);
                }
            });

            const imageInput = document.querySelector('input[name="image"]');
            imageInput.addEventListener('change', showFileName);

            htmlEditor.getSession().on('change', updatePreview);
            cssEditor.getSession().on('change', updatePreview);
        });
    </script>
</head>
<body>
    <h1>Excel to HTML Table Converter</h1>
    <form method="POST" enctype="multipart/form-data">
        <textarea name="excel_data" rows="10" cols="50" placeholder="Paste your Excel data here"></textarea><br>
        <button type="submit" name="action" value="convert">Convert</button>
    </form>
    {% if table %}
    <div class="editor-container">
        <div class="editor-section">
            <h2>HTML Editor</h2>
            <div id="editor">{{ table }}</div>
            <textarea id="html_code" name="html_code" style="display:none;">{{ table }}</textarea>
            <h2>CSS Editor</h2>
            <div id="css-editor">/* Enter your CSS here */</div>
        </div>
        <div class="preview-section">
            <h2>Preview</h2>
            <iframe id="preview"></iframe>
        </div>
    </div>
    {% endif %}

    <h1>Image Compressor</h1>
    <form onsubmit="uploadImage(event)">
        <label class="file-upload-label" for="file-upload">Choose File</label>
        <input id="file-upload" class="file-upload" type="file" name="image" accept="image/*"><br>
        <span id="file_info" class="file-info"></span><br>
        <button type="submit">Upload and Compress</button>
    </form>
    <div id="compressed_image_section" style="display:none;">
        <h2>Compressed Image</h2>
        <a id="compressed_image_link" class="download-link" href="#" download>Download Compressed Image</a>
        <div id="size_info" class="size-info"></div>
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
                original_size = image.content_length
                compressed_image = compress_image(image)
                compressed_image_url = url_for('download_file', filename=compressed_image, _external=True)
                compressed_image_size = os.path.getsize(os.path.join(app.config['UPLOAD_FOLDER'], compressed_image))
                return jsonify({
                    'compressed_image': compressed_image,
                    'compressed_image_url': compressed_image_url,
                    'compressed_image_size': compressed_image_size
                })

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
