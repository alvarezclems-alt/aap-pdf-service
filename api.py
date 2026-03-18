"""
api.py — Serveur Flask pour la génération des Annexes PDF officielles INSPÉ
Appel depuis l'app Lovable via POST /generate-annexe1 et /generate-annexe1bis

Démarrage : python api.py
Port par défaut : 5050
"""
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import io, os, traceback, zipfile
from generate_annexes import build_annexe1, build_annexe1bis

app = Flask(__name__)
CORS(app)

LOGO_PATH = os.path.join(os.path.dirname(__file__), 'logo_inspe.png')

def get_logo():
    try:
        with open(LOGO_PATH, 'rb') as f:
            return f.read()
    except:
        return None

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'service': 'AAP PDF Generator'})

@app.route('/generate-annexe1', methods=['POST'])
def route_annexe1():
    try:
        data = request.get_json(force=True)
        pdf_bytes = build_annexe1(data.get('project', {}), get_logo())
        acronyme = data.get('project', {}).get('acronyme', 'projet').replace(' ', '_')
        return send_file(io.BytesIO(pdf_bytes), mimetype='application/pdf',
                         as_attachment=True, download_name=f'Annexe1_{acronyme}.pdf')
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/generate-annexe1bis', methods=['POST'])
def route_annexe1bis():
    try:
        data = request.get_json(force=True)
        pdf_bytes = build_annexe1bis(data.get('project', {}),
                                     data.get('budget_lines', []), get_logo())
        titre = data.get('project', {}).get('titre', 'budget')[:30].replace(' ', '_')
        return send_file(io.BytesIO(pdf_bytes), mimetype='application/pdf',
                         as_attachment=True, download_name=f'Annexe1bis_{titre}.pdf')
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/generate-all', methods=['POST'])
def route_all():
    """Génère les deux PDFs et les retourne dans un ZIP."""
    try:
        data     = request.get_json(force=True)
        project  = data.get('project', {})
        lines    = data.get('budget_lines', [])
        logo     = get_logo()
        acronyme = project.get('acronyme', 'projet').replace(' ', '_')

        pdf1 = build_annexe1(project, logo)
        pdf2 = build_annexe1bis(
            {'titre':         project.get('titre', ''),
             'porteur':       project.get('coordinateur', {}).get('nom', ''),
             'montant_inspe': project.get('montant_inspe', 8000)},
            lines, logo)

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
            zf.writestr(f'Annexe1_{acronyme}.pdf',    pdf1)
            zf.writestr(f'Annexe1bis_{acronyme}.pdf', pdf2)
        buf.seek(0)

        return send_file(buf, mimetype='application/zip',
                         as_attachment=True,
                         download_name=f'Dossier_AAP_{acronyme}.zip')
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5050))
    print(f'AAP PDF Service → http://localhost:{port}')
    app.run(host='0.0.0.0', port=port, debug=False)
