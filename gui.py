import tkinter as tk
from main import speak, listen, run_command

class JarvisGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Jarvis Assistant")
        self.root.geometry("600x500")
        self.root.config(bg="#20232A")

        self.text_box = tk.Text(root, wrap="word", bg="#282C34", fg="#61DAFB", font=("Consolas", 12))
        self.text_box.pack(padx=10, pady=10, fill="both", expand=True)

        self.listen_button = tk.Button(
            root,
            text="🎤 Speak",
            command=self.process_voice,
            bg="#61DAFB",
            fg="#20232A",
            font=("Arial", 12, "bold")
        )
        self.listen_button.pack(pady=10)

    def process_voice(self):
        command = listen()
        if not command:
            return
        self.text_box.insert(tk.END, f"\n🧍You: {command}\n")
        response = run_command(command)
        if response:
            self.text_box.insert(tk.END, f"🤖 Jarvis: {response}\n")
        self.text_box.see(tk.END)

if __name__ == "__main__":
    root = tk.Tk()
    app = JarvisGUI(root)
    root.mainloop()
