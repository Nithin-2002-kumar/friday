from main import speak
from gui import JarvisGUI
import tkinter as tk

if __name__ == "__main__":
    speak("Jarvis is online and ready.")
    root = tk.Tk()
    app = JarvisGUI(root)
    root.mainloop()
