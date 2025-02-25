import tkinter as tk
from tkinter import messagebox

class CountdownApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Choose an Option")
        
        self.label = tk.Label(root, text="Choose an option before time runs out!", font=("Arial", 12))
        self.label.pack(pady=10)
        
        self.countdown_label = tk.Label(root, text="5", font=("Arial", 14, "bold"))
        self.countdown_label.pack()
        
        self.option1_button = tk.Button(root, text="Option 1 (Default)", command=self.select_option1)
        self.option1_button.pack(pady=5)
        
        self.option2_button = tk.Button(root, text="Option 2", command=self.select_option2)
        self.option2_button.pack(pady=5)
        
        self.time_left = 5
        self.update_countdown()
    
    def update_countdown(self):
        if self.time_left > 0:
            self.countdown_label.config(text=str(self.time_left))
            self.time_left -= 1
            self.root.after(1000, self.update_countdown)
        else:
            self.select_option1()
    
    def select_option1(self):
        messagebox.showinfo("Selection", "You selected Option 1")
        self.root.quit()
    
    def select_option2(self):
        messagebox.showinfo("Selection", "You selected Option 2")
        self.root.quit()

if __name__ == "__main__":
    root = tk.Tk()
    app = CountdownApp(root)
    root.mainloop()
