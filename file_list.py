from pathlib import Path


def list_video_files_in_directory(directory):
    video_extensions = ['.mp4', '.avi', '.mov', '.mkv', '.wmv']
    path = Path('uploads')
    fileList = []
    fileData = {}
    if path.exists() and path.is_dir():
        video_files = [file for file in path.iterdir() if file.suffix.lower() in video_extensions]
        for video in video_files:
            print(video.name + " :: " + str(video.resolve()))
            fileData['name'] = video.name
            fileData['path'] = str(video.resolve())
            fileData['time'] = video.stat().st_mtime
            fileData['size'] = video.stat().st_size
            fileData['created'] = video.stat().st_ctime
            fileList.append(fileData)
            
    
    else:
        print(f"The directory {directory} does not exist or is not a directory.")
        
    print(fileList)

# 사용 예시
list_video_files_in_directory('uploads')
