from pathlib import Path


def list_video_files_in_directory(directory):
    video_extensions = ['.mp4', '.avi', '.mov', '.mkv', '.wmv']
    path = Path(directory)
    if path.exists() and path.is_dir():
        video_files = [file for file in path.iterdir() if file.suffix.lower() in video_extensions]
        for video in video_files:
            print(video.name + " :: " + str(video.resolve()))
    else:
        print(f"The directory {directory} does not exist or is not a directory.")

# 사용 예시
list_video_files_in_directory('uploads')
