import os
import zipfile
from flask import Flask, request, render_template
from flask_mail import Mail, Message
from lxml import etree
from dotenv import load_dotenv

# Load credentials from .env file
load_dotenv()

app = Flask(__name__)

# Configuration
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAIL_SERVER'] = 'smtp.gmail.com' # Change if using UCL Outlook
app.config['MAIL_PORT'] = 587
app.config['MAIL_USE_TLS'] = True
app.config['MAIL_USERNAME'] = os.getenv('MAIL_USERNAME')
app.config['MAIL_PASSWORD'] = os.getenv('MAIL_PASSWORD')
app.config['MAIL_DEFAULT_SENDER'] = os.getenv('MAIL_SENDER')

mail = Mail(app)

# Ensure upload folder exists
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

def check_for_track_changes(file_path):
    """Deep scan of .docx XML structure for track changes."""
    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            # Check 1: Existing Revisions (ins/del tags)
            doc_xml = z.read('word/document.xml')
            tree = etree.fromstring(doc_xml)
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            revisions = tree.xpath('//w:ins | //w:del', namespaces=ns)
            if revisions:
                return True, "Unaccepted revisions found."

            # Check 2: Track Changes Mode Toggle
            if 'word/settings.xml' in z.namelist():
                settings_xml = z.read('word/settings.xml')
                settings_tree = etree.fromstring(settings_xml)
                track_on = settings_tree.xpath('//w:trackRevisions', namespaces=ns)
                if track_on:
                    return True, "Track Changes mode is active."
                    
        return False, "Clean"
    except Exception as e:
        return True, f"Parsing error: {str(e)}"

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return "No file provided", 400
    
    file = request.files['file']
    if file.filename == '':
        return "No file selected", 400

    filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(filepath)

    is_rejected, reason = check_for_track_changes(filepath)

    if is_rejected:
        os.remove(filepath)
        return f"Rejected: {reason}", 403

    # Send Email
    try:
        msg = Message("Approved Document Uploaded",
                      recipients=["g.dash@ucl.ac.uk"])
        msg.body = f"The attached document '{file.filename}' passed the check."
        with app.open_resource(filepath) as fp:
            msg.attach(file.filename, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", fp.read())
        
        mail.send(msg)
        os.remove(filepath) # Clean up after sending
        return "Document approved and sent to g.dash@ucl.ac.uk"
    except Exception as e:
        return f"Error sending email: {str(e)}", 500

if __name__ == '__main__':
    app.run(debug=True)
