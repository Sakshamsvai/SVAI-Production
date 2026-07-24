from flask import Blueprint, request, jsonify

uploads_bp = Blueprint('uploads', __name__)

@uploads_bp.route('/api/upload-valuation-folder', methods=['POST'])
def upload_folder():
    files = request.files.getlist('files')
    return jsonify({"success": True, "message": f"Valuation Folder ({len(files)} files) Upload Ho Gaya!"})

@uploads_bp.route('/api/upload-format', methods=['POST'])
def upload_format():
    file = request.files.get('file')
    filename = file.filename if file else 'Format'
    return jsonify({"success": True, "message": f"Format File ({filename}) Upload Ho Gayi!"})

@uploads_bp.route('/api/upload-visit-zip', methods=['POST'])
def upload_zip():
    file = request.files.get('file')
    filename = file.filename if file else 'Visit Data'
    return jsonify({"success": True, "message": f"Visit Data ZIP ({filename}) Upload Ho Gaya!"})
