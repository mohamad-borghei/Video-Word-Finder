import os
from moviepy.editor import VideoFileClip


def extract_audio_from_video(video_path, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    video = VideoFileClip(video_path)
    audio = video.audio

    audio_filename = os.path.splitext(os.path.basename(video_path))[0] + ".wav"
    audio_output_path = os.path.join(output_folder, audio_filename)

    audio.write_audiofile(audio_output_path, codec='pcm_s16le')

    print(f"The audio file was saved in the following path: {audio_output_path}")
    return audio_output_path
