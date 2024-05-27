import ctypes
import json
import os
from ctypes import windll
from datetime import datetime
from threading import Thread
from time import sleep

from moviepy.editor import ImageClip, concatenate_videoclips, VideoFileClip, CompositeVideoClip
from moviepy.video.fx.fadein import fadein
from moviepy.video.fx.fadeout import fadeout
from moviepy.video.fx.resize import resize


# Global variables
VIDEO_DURATION = 1
ROOT = "\\".join(os.path.abspath(__file__).split('\\')[:-1])
user32 = ctypes.windll.user32
WINDOW_SIZE = (user32.GetSystemMetrics(0), user32.GetSystemMetrics(1))
clip_duration: int
schedule = json.load(open(ROOT+'\\data\\schedule.json', 'r'))


# Global functions

def get_center_pos(size: tuple[int | float]) -> tuple[float]:
    return (WINDOW_SIZE[0] / 2) - (size[0] / 2), (WINDOW_SIZE[1] / 2) - (size[1] / 2)


def get_correct_size(size_before: tuple[int | float]) -> tuple[float]:
    k0 = WINDOW_SIZE[0] / size_before[0]
    if k0 * size_before[1] > WINDOW_SIZE[1]:
        k1 = WINDOW_SIZE[1] / size_before[1]
        return k1 * size_before[0], k1 * size_before[1]
    return k0 * size_before[0], k0 * size_before[1]


# Threads

def slide_show_manager():
    get_slide_show()
    play_slide_show()


def get_slide_show():
    global clip_duration

    # Get all media from the media directory
    media_group = os.listdir(ROOT + '\\media\\ads')

    # Containers with videos and photos
    video_clips: list[VideoFileClip | CompositeVideoClip] = []
    img_clips: list[ImageClip] = []

    # Get clips of media files, and sort it by containers
    for media in media_group:
        package = ROOT + f'\\media\\ads\\{media}'
        if media.split('.')[-1].lower() in ["avi", "mp4", "webm"]:
            vid = VideoFileClip(package)
            size = get_correct_size(vid.size)
            vid = resize(vid, width=size[0], height=size[1])
            video_clips.append(vid)
        else:
            img = ImageClip(package)
            size = get_correct_size(img.size)
            img = resize(img, width=size[0], height=size[1])
            img = img.set_duration(3)
            img_clips.append(img)

    """
    if video count is not zero,
    then distribute photos among videos,
    create video clips from image clips and place them between videos
    """
    if not len(video_clips) == 0:
        video_count = len(img_clips) // len(video_clips)
        for i in range(0, len(video_clips) - 1):
            video_clips.insert(
                2 * i + 1,
                concatenate_videoclips(
                    img_clips[i * video_count:(i + 1) * video_count],
                    method="compose"
                )
            )
            remainder = len(img_clips) - len(video_clips) * video_count
            video_clips += [concatenate_videoclips(img_clips[-remainder:], method="compose")]

    # Set to every video clip fade in and fade out effects, and centering it
    for i, video_clip in enumerate(video_clips):
        video_clips[i] = (video_clip.fx(fadein, 1)
                                    .fx(fadeout, 1)
                                    .set_position(get_center_pos(video_clip.size))
                          )

    # Compose all video clips, and set them start time in final clip by algorithm. After write it to file
    final_clip = CompositeVideoClip(
        [video_clips[0]] + [
            video_clip.set_start(
                sum([video_clips[i].duration for i in range(0, i+1)])
            ).crossfadein(1)
            for i, video_clip in enumerate(video_clips[1:])
        ],
        size=WINDOW_SIZE
    )
    final_clip.write_videofile(ROOT + f'\\media\\slide_show\\slide_show.mp4', fps=30)

    # Set global variable clip_duration with final clip duration
    clip_duration = final_clip.duration


def play_slide_show():
    path = ROOT + "\\media\\slide_show\\slide_show.mp4"
    while True:
        os.system(f'start {path} /max slide_show.mp4')
        sleep(clip_duration)


def shutdown():
    """
    Thread that control shutdown of Windows system

    :return:
    """
    # Get shutdown datetime
    now = datetime.now()
    shutdown_time = schedule[now.weekday()].split(':')
    shutdown_date = datetime(
        year=now.year,
        month=now.month,
        day=now.day,
        hour=int(shutdown_time[0]),
        minute=int(shutdown_time[1])
    )

    # Wait for shutdown
    delta = shutdown_date - now
    sleep(delta.seconds)

    # Shutdown
    bat_file_path = ROOT + "\\shutdown.bat"  # from OP
    result = windll.shell32.ShellExecuteW(
        None,  # handle to parent window
        'runas',  # verb
        'cmd.exe',  # file on which verb acts
        ' '.join(['/c', bat_file_path]),  # parameters
        None,  # working directory (default is cwd)
        1,  # show window normally
    )
    return result > 32


# Main

def main():
    slide_show_thread = Thread(target=slide_show_manager)
    shutdown_thread = Thread(target=shutdown)
    slide_show_thread.start()
    shutdown_thread.start()


if __name__ == "__main__":
    main()
