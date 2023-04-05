import time

class Project:

    def __init__(self):
        self.project_name = "project"
        self.project_time = "time"
        self.is_active = False
        self.cancel_var = False
        self.project_pause = False

    def start_time(self, inc = 0):
        self.time_started = time.time() - inc


    def stop_time(self):
        self.time_ended = time.time()
        self.delta_time = self.time_ended-self.time_started
        self.project_time = float(self.delta_time / 3600) #round(float(self.delta_time / 3600),1)
        print(self.delta_time)
        return self.project_time


    def cancel(self):
        self.delta_time = 0
        self.project_time = 0
        self.cancel_var = True
