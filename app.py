from flask import Flask, send_file, jsonify
import os

app = Flask(__name__)

@app.route('/')
def index():
    return send_file('index.html')

@app.route('/thumbnails/<path:filename>')
def thumbnails(filename):
    path = os.path.join('thumbnails', filename)
    if os.path.exists(path):
        return send_file(path)
    else:
        return '', 404

@app.route('/get_thumbs/<path:rel_path>')
def get_thumbs(rel_path):
    base = os.path.join('thumbnails', rel_path)
    thumbs = []
    i = 1
    while os.path.exists(f'{base}_{i}.jpg'):
        thumbs.append(f'/thumbnails/{rel_path}_{i}.jpg')
        i += 1
    if thumbs:
        return jsonify(thumbs)
    else:
        return jsonify([])

if __name__ == '__main__':
    app.run(debug=True)