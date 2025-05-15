# import os
# import tempfile
# from flask import Flask, request, redirect, url_for, render_template, send_file, flash
# import pythoncom
# import win32com.client

# app = Flask(__name__)
# app.secret_key = 'replace-with-a-secure-random-key'

# # Ensure upload folder exists
# UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), 'uploads')
# os.makedirs(UPLOAD_FOLDER, exist_ok=True)
# app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# def convert_ppt_to_pptx(ppt_path: str) -> str:
#     pythoncom.CoInitialize()
#     try:
#         ppt_app = win32com.client.Dispatch("PowerPoint.Application")
#         # Leave ppt_app.Visible at its default (True)
#         # Open with no window:
#         pres = ppt_app.Presentations.Open(ppt_path,
#                                            WithWindow=False)
#         # (Optional) minimize if you see any UI flicker:
#         ppt_app.WindowState = 2  # ppWindowMinimized

#         fd, out_path = tempfile.mkstemp(suffix='.pptx',
#                                          dir=UPLOAD_FOLDER)
#         os.close(fd)
#         pres.SaveAs(out_path, 24)  # ppSaveAsOpenXMLPresentation
#         pres.Close()
#         return out_path
#     finally:
#         ppt_app.Quit()
#         pythoncom.CoUninitialize()


# @app.route('/', methods=['GET', 'POST'])
# def index():
#     if request.method == 'POST':
#         uploaded = request.files.get('file')
#         if not uploaded or not uploaded.filename.lower().endswith('.ppt'):
#             flash('Please upload a valid .ppt file.')
#             return redirect(request.url)

#         # Save the .ppt to a temp file
#         in_fd, in_path = tempfile.mkstemp(suffix='.ppt', dir=UPLOAD_FOLDER)
#         os.close(in_fd)
#         uploaded.save(in_path)

#         # Convert and clean up input
#         try:
#             out_path = convert_ppt_to_pptx(in_path)
#         except Exception as e:
#             flash(f'Conversion failed: {e}')
#             os.remove(in_path)
#             return redirect(request.url)

#         # Remove the original .ppt
#         os.remove(in_path)

#         # Redirect to download page
#         return redirect(url_for('result', fname=os.path.basename(out_path)))

#     return render_template('index.html')

# @app.route('/result/<fname>')
# def result(fname):
#     out_path = os.path.join(app.config['UPLOAD_FOLDER'], fname)
#     if not os.path.exists(out_path):
#         flash('File not found.')
#         return redirect(url_for('index'))
#     return render_template('result.html', filename=fname)

# @app.route('/download/<filename>')
# def download(filename):
#     path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
#     return send_file(path, as_attachment=True)

# if __name__ == '__main__':
#     app.run(debug=True)
import io
import os
from flask import (
    Flask, request, redirect, url_for,
    render_template, send_file, flash
)
import aspose.slides as slides
import aspose.pydrawing as drawing  # required on some platforms

app = Flask(__name__)
app.secret_key = 'replace-with-a-secure-random-key'

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        uploaded = request.files.get('file')
        if not uploaded or not uploaded.filename.lower().endswith('.ppt'):
            flash('Please upload a valid .ppt file.')
            return redirect(request.url)

        # 1) Read the uploaded PPT into memory
        in_bytes = uploaded.read()
        in_stream = io.BytesIO(in_bytes)

        # 2) Convert PPT → PPTX in-memory
        out_stream = io.BytesIO()
        with slides.Presentation(in_stream) as pres:           # :contentReference[oaicite:2]{index=2}
            pres.save(out_stream, slides.export.SaveFormat.PPTX) # :contentReference[oaicite:3]{index=3}

        # 3) Stream the result back to the user
        out_stream.seek(0)
        return send_file(
            out_stream,
            as_attachment=True,
            download_name=os.path.splitext(uploaded.filename)[0] + '.pptx',
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
