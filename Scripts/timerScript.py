__author__ = 'Rajiv'
import winsound
import time

minutes = 19
seconds = 40


def count_down(inp_minutes, inp_seconds):
    total_seconds = inp_minutes*60+inp_seconds
    print(total_seconds)
    for each_second in range(total_seconds):
        minutes_remaining = (total_seconds-each_second) // 60
        seconds_remaining = (total_seconds-each_second) % 60
        print("time remaining: "+str(minutes_remaining)+":"+str(seconds_remaining).zfill(2))
        time.sleep(1)
    winsound.Beep(440,1000)

count_down(minutes,seconds)
