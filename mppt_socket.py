import os
import urllib.parse

from flask import Flask, jsonify
from flask_socketio import SocketIO, emit

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
socketio = SocketIO(app, cors_allowed_origins="*")

if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

@socketio.on('connect')
def handle_connect():
    print('Client connected')

@socketio.on('disconnect')
def handle_disconnect():
    print('Client disconnected')

@socketio.on('upload_video_chunk')
def handle_upload_video_chunk(data):
    # 데이터에서 필요한 정보 추출
    file_name = data['file_name']
    file_name = urllib.parse.unquote(file_name)
    chunk_data = data['chunk_data']
    chunk_number = data['chunk_number']
    total_chunks = data['total_chunks']
    
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], file_name)
    
    # 청크 데이터를 파일에 추가 (ab 모드로 파일을 열어 데이터 추가)
    with open(file_path, 'ab') as f:
        f.write(chunk_data)
    
    if chunk_number == total_chunks:
        # 모든 청크가 전송된 경우
        emit('upload_complete', {"file_path": file_path})
    else:
        emit('chunk_received', {"current" : chunk_number, "total": total_chunks})

if __name__ == '__main__':
    socketio.run(app,host='0.0.0.0',port=5003, debug=False)
