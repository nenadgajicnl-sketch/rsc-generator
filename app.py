from flask import Flask, request, jsonify
from flask_cors import CORS
import subprocess, tempfile, os, base64

app = Flask(__name__)
CORS(app)

@app.route('/convert', methods=['POST'])
def convert():
    try:
        data = request.get_json()
        docx_b64 = data.get('docx')
        filename = data.get('filename', 'document')
        
        docx_bytes = base64.b64decode(docx_b64)
        
        with tempfile.TemporaryDirectory() as tmpdir:
            docx_path = os.path.join(tmpdir, f'{filename}.docx')
            pdf_path = os.path.join(tmpdir, f'{filename}.pdf')
            
            with open(docx_path, 'wb') as f:
                f.write(docx_bytes)
            
            result = subprocess.run([
                'libreoffice', '--headless', '--convert-to', 'pdf',
                '--outdir', tmpdir, docx_path
            ], capture_output=True, text=True, timeout=60)
            
            if not os.path.exists(pdf_path):
                return jsonify({'error': result.stderr or result.stdout}), 500
            
            with open(pdf_path, 'rb') as f:
                pdf_b64 = base64.b64encode(f.read()).decode()
            
            return jsonify({'pdf': pdf_b64})
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/', methods=['GET'])
@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port, debug=False)
