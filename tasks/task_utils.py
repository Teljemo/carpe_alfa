from datetime import datetime, timedelta

class Task:
    def __init__(self, name):
        self.name = name
        self.start_time = None
        self.pause_time = None
        self.elapsed = timedelta()
        self.active = False
        self.paused = False

    def start(self):
        if not self.active:
            self.start_time = datetime.now()
            self.active = True
            self.paused = False
        elif self.paused:
            # återuppta från paus
            delta = datetime.now() - self.pause_time
            self.start_time += delta
            self.paused = False

    def pause(self):
        if self.active and not self.paused:
            self.pause_time = datetime.now()
            self.elapsed += self.pause_time - self.start_time
            self.paused = True

    def stop(self):
        if self.active:
            if not self.paused:
                self.elapsed += datetime.now() - self.start_time
            self.active = False
            self.paused = False
            self.start_time = None
            self.pause_time = None

    def get_elapsed_seconds(self):
        if self.active and not self.paused:
            return (datetime.now() - self.start_time + self.elapsed).total_seconds()
        return self.elapsed.total_seconds()

    def get_elapsed_str(self):
        total = int(self.get_elapsed_seconds())
        h, remainder = divmod(total, 3600)
        m, s = divmod(remainder, 60)
        return f"{h:02d}:{m:02d}:{s:02d}"
