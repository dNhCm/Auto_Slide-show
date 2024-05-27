
import ctypes
import json
import os
from ctypes import windll
from datetime import datetime

from PIL import Image as Img

from kivy.app import App
from kivy.clock import Clock
from kivy.core.window import Window
from kivy.graphics import Color
from kivy.properties import partial
from kivy.uix.image import Image
from kivy.uix.video import Video
from kivy.uix.widget import Widget
from moviepy.editor import VideoFileClip


# Global variables
VIDEO_DURATION = 7
ROOT = "\\".join(os.path.abspath(__file__).split('\\')[:-1])
user32 = ctypes.windll.user32
WINDOW_SIZE = (user32.GetSystemMetrics(0), user32.GetSystemMetrics(1))


class SlideShow(Widget):
    def __init__(self):
        super().__init__()
        self.size = Window.size
        with self.canvas:
            Color(0, 0, 0, 0)
        self.i = 0

        Clock.schedule_once(partial(self.update, self))

    @staticmethod
    def get_center_pos(size: tuple[int | float]) -> tuple[float]:
        return (WINDOW_SIZE[0] / 2) - (size[0] / 2), (WINDOW_SIZE[1] / 2) - (size[1] / 2)

    @staticmethod
    def get_correct_size(size_before: tuple[int | float]) -> tuple[float]:
        k0 = WINDOW_SIZE[0] / size_before[0]
        if k0*size_before[1] > WINDOW_SIZE[1]:
            k1 = WINDOW_SIZE[1] / size_before[1]
            return k1*size_before[0], k1*size_before[1]
        return k0*size_before[0], k0*size_before[1]

    def update(self, *args, **kwargs):
        """
        Update task for slide showing

        :param args:
        :param kwargs:
        :return:
        """
        # Get all photos from media
        media_group = os.listdir(ROOT+'/media')

        # Get current photo, if it's range out of our photos array, then start it again
        try:
            media = media_group[self.i]
        except IndexError:
            media = media_group[0]
            self.i = 0

        # Get package of media (abspath)
        package = ROOT+f"\\media\\{media}"

        # Return photo or video (.avi only!) and count duration of video to wait for next media
        video_duration = VIDEO_DURATION
        if media.split('.')[-1] in ["avi"]:
            vid = VideoFileClip(package)
            size = self.get_correct_size(vid.size)
            video_duration = vid.duration

            self.canvas.after.clear()
            with self.canvas.after:
                Video(
                    source=package,
                    size=size,
                    pos=self.get_center_pos(size),
                    allow_stretch=True,
                    play=True
                )
        else:
            img = Img.open(package)
            size = SlideShow.get_correct_size(img.size)

            self.canvas.after.clear()
            with self.canvas.after:
                Image(
                    source=package,
                    size=size,
                    pos=SlideShow.get_center_pos(size),
                    allow_stretch=True
                )

        # Add one to our index line for next media, and start again our update task after vide_duration time
        self.i += 1
        Clock.schedule_once(partial(self.update, self), video_duration)


class SlideShowApp(App):
    @staticmethod
    def get_shutdown_delta() -> int:
        """
        Count how many seconds need to wait before shutdown

        :return: count of seconds to wait
        """
        # Get schedule
        with open(ROOT+'/data/schedule.json', 'r') as jsonfile:
            schedule = json.load(jsonfile)

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

        return delta.seconds

    @staticmethod
    def shutdown(*arg):
        """
        Task that shutdown of Windows system

        :return:
        """
        bat_file_path = ROOT+"\\shutdown.bat"  # from OP
        result = windll.shell32.ShellExecuteW(
            None,  # handle to parent window
            'runas',  # verb
            'cmd.exe',  # file on which verb acts
            ' '.join(['/c', bat_file_path]),  # parameters
            None,  # working directory (default is cwd)
            1,  # show window normally
        )
        return result > 32

    def build(self):
        # Window presets
        Window.size = WINDOW_SIZE
        Window.fullscreen = "auto"

        # Get slide-show
        slide_show = SlideShow()

        Clock.schedule_once(self.shutdown, self.get_shutdown_delta())

        # Return slide-show
        return slide_show


if __name__ == "__main__":
    app = SlideShowApp()
    try:
        app.run()
    except (KeyboardInterrupt, SystemExit):
        app.stop()
