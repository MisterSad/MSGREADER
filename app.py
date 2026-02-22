# -*- coding: utf-8 -*-
# --------------------------------------------------
# MSG READER pour macOS
# Copyright (C) 2025 Marko8787
#
# Ce programme est un logiciel libre ; vous pouvez le redistribuer et/ou le
# modifier selon les termes de la Licence Publique Générale GNU telle que
# publiée par la Free Software Foundation ; soit la version 3 de la
# Licence, ou (à votre choix) toute version ultérieure.
#
# Ce programme est distribué dans l'espoir qu'il sera utile, mais SANS
# AUCUNE GARANTIE ; sans même la garantie implicite de COMMERCIALISATION
# ou D'ADAPTATION À UN USAGE PARTICULIER. Consultez la Licence Publique
# Générale GNU pour plus de détails.
#
# Vous devriez avoir reçu une copie de la Licence Publique Générale GNU
# avec ce programme ; si ce n'est pas le cas, consultez
# <https://www.gnu.org/licenses/>.
#
# --------------------------------------------------
# Version : 1.2 (Licence et finalisation)
# --------------------------------------------------

# --- Importation des modules nécessaires ---
import os
import sys
import shutil
import webview
import extract_msg
import subprocess
from flask import Flask, request, render_template, jsonify, url_for
from werkzeug.utils import secure_filename
from bs4 import BeautifulSoup
from email.message import EmailMessage
from email.utils import formatdate, make_msgid
import mimetypes

# --- DÉFINITION DES CHEMINS ---
def get_base_path():
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    else:
        return os.path.dirname(os.path.abspath(__file__))

WORK_DIR = os.path.join(os.path.expanduser('~'), 'MSGREADER_FichiersTemporaires')
TEMPLATE_FOLDER = os.path.join(get_base_path(), 'templates')
STATIC_FOLDER = os.path.join(get_base_path(), 'static')

# --- CONFIGURATION DE L'APPLICATION FLASK ---
app = Flask(__name__, template_folder=TEMPLATE_FOLDER, static_folder=STATIC_FOLDER)
app.secret_key = 'la_cle_de_la_version_github'
app.config['WORK_FOLDER'] = WORK_DIR
os.makedirs(app.config['WORK_FOLDER'], exist_ok=True)

# ----- API DE COMMUNICATION ENTRE PYTHON ET JAVASCRIPT -----
class Api:
    def __init__(self):
        self.window = None
        self.current_msg_path = None

    def open_attachment(self, filename):
        safe_filename = secure_filename(filename)
        if not safe_filename: return {'status': 'error', 'message': 'Nom de fichier invalide.'}
        file_path = os.path.join(app.config['WORK_FOLDER'], safe_filename)
        if os.path.exists(file_path):
            try: subprocess.run(['open', file_path], check=True)
            except Exception as e: return {'status': 'error', 'message': str(e)}
        return {'status': 'error', 'message': 'Fichier non trouvé.'}

    def open_external_link(self, url):
        if url.startswith(('http://', 'https://', 'mailto:')):
            try: subprocess.run(['open', url], check=True)
            except Exception as e: return {'status': 'error', 'message': str(e)}
        else:
            return {'status': 'error', 'message': 'Lien externe non autorisé.'}

    def reveal_attachments(self):
        try: subprocess.run(['open', app.config['WORK_FOLDER']], check=True)
        except Exception as e: return {'status': 'error', 'message': str(e)}

    def export_to_eml(self):
        if not self.current_msg_path or not os.path.exists(self.current_msg_path):
            return {'status': 'error', 'message': 'Fichier .msg original non trouvé.'}

        try:
            msg_data = extract_msg.Message(self.current_msg_path)
            eml = EmailMessage()

            if msg_data.subject: eml['Subject'] = msg_data.subject
            if msg_data.sender: eml['From'] = msg_data.sender
            if msg_data.to: eml['To'] = msg_data.to
            if msg_data.cc: eml['Cc'] = msg_data.cc
            if msg_data.date: eml['Date'] = msg_data.date
            else: eml['Date'] = formatdate(localtime=True)
            eml['Message-ID'] = make_msgid()

            if msg_data.htmlBody:
                eml.set_content("Ce message nécessite un client mail compatible HTML pour être affiché correctement.", subtype='plain')
                eml.add_alternative(msg_data.htmlBody.decode('utf-8', 'ignore'), subtype='html')
            elif msg_data.body:
                eml.set_content(msg_data.body)
            else:
                 eml.set_content("(Ce message n'a pas de contenu)")

            if msg_data.attachments:
                for attachment in msg_data.attachments:
                    try:
                        ctype, _ = mimetypes.guess_type(attachment.longFilename)
                        if ctype is None: ctype = 'application/octet-stream'
                        maintype, subtype = ctype.split('/', 1)
                        eml.add_attachment(attachment.data, maintype=maintype, subtype=subtype, filename=os.path.basename(attachment.longFilename))
                    except Exception:
                        continue

            eml_filename = os.path.splitext(os.path.basename(self.current_msg_path))[0] + ".eml"
            eml_path = os.path.join(app.config['WORK_FOLDER'], eml_filename)
            with open(eml_path, 'wb') as f: f.write(eml.as_bytes())
            
            subprocess.run(['open', eml_path], check=True)
            return {'status': 'success'}
        except Exception as e:
            return {'status': 'error', 'message': f"Erreur d'export : {e}"}

    def set_window_title(self, title):
        if self.window: self.window.set_title(title)

# --- ROUTES DE L'APPLICATION WEB ---
@app.route('/')
def index():
    icon_url = url_for('static', filename='icon.png')
    return render_template('index.html', icon_url=icon_url)

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files: return jsonify({'status': 'error', 'message': 'Aucun fichier reçu.'})
    file = request.files['file']
    if file.filename == '': return jsonify({'status': 'error', 'message': 'Aucun fichier sélectionné.'})

    if file and file.filename.lower().endswith('.msg'):
        shutil.rmtree(app.config['WORK_FOLDER'])
        os.makedirs(app.config['WORK_FOLDER'], exist_ok=True)
        filename = secure_filename(file.filename)
        msg_path = os.path.join(app.config['WORK_FOLDER'], filename)
        file.save(msg_path)
        api.current_msg_path = msg_path

        try: msg = extract_msg.Message(msg_path)
        except Exception as e: return jsonify({'status': 'error', 'message': f'Fichier .msg invalide ou corrompu : {e}'})

        try: subject = msg.subject
        except: subject = "(Sujet illisible)"
        try: sender = msg.sender
        except: sender = "(Expéditeur inconnu)"
        try: to = msg.to
        except: to = "(Destinataire inconnu)"
        
        email_body = ""
        attachments_list = []
        cid_map = {}

        if msg.attachments:
            for attachment in msg.attachments:
                try:
                    sanitized_filename = os.path.basename(attachment.longFilename)
                    full_save_path = os.path.join(app.config['WORK_FOLDER'], sanitized_filename)
                    with open(full_save_path, 'wb') as f: f.write(attachment.data)
                    attachments_list.append(sanitized_filename)
                    if attachment.cid: cid_map[attachment.cid] = full_save_path
                except: continue
        
        try:
            if msg.htmlBody:
                soup = BeautifulSoup(msg.htmlBody.decode('utf-8', 'ignore'), 'html.parser')
                
                # Sanitize HTML tags
                for script in soup(['script', 'iframe', 'object', 'embed', 'applet']):
                    script.decompose()

                # Sanitize dangerous attributes
                for tag in soup.find_all(True):
                    attrs_to_remove = [attr for attr in tag.attrs if attr.lower().startswith('on') or attr.lower() == 'formaction']
                    for attr in attrs_to_remove:
                        del tag[attr]

                for img in soup.find_all('img'):
                    src = img.get('src')
                    if src and src.startswith('cid:'):
                        cid = src[4:]
                        if cid in cid_map: img['src'] = 'file://' + cid_map[cid]
                email_body = str(soup)
            elif msg.body:
                email_body = f"<pre style='white-space: pre-wrap; word-wrap: break-word; font-family: sans-serif;'>{msg.body}</pre>"
            else:
                email_body = "<i>(Ce message n'a pas de contenu visible.)</i>"
        except Exception as e:
            email_body = f"<p style='color:red;'><b>Impossible de lire le corps du message.</b><br>Erreur : {e}</p>"
        
        return jsonify({ 'status': 'success', 'subject': subject, 'from': sender, 'to': to, 'body': email_body, 'attachments': attachments_list })
    
    return jsonify({'status': 'error', 'message': 'Type de fichier non valide.'})

# ----- Lancement de l'application -----
if __name__ == '__main__':
    api = Api()
    window = webview.create_window('MSG READER', app, js_api=api, width=800, height=700, resizable=True)
    api.window = window
    webview.start()