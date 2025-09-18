import pygame
pygame.mixer.init()
try:
    pygame.mixer.music.load("Moon.mp3")
    pygame.mixer.music.play()
    input("播放中，按 Enter 結束…")
except Exception as e:
    print("Error:", e)
