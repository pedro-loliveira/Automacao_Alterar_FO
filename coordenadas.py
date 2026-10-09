import tkinter as tk
import pyautogui
import time


class MousePositionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Captura de Coordenadas do Mouse")

        self.label = tk.Label(
            root, text="Mover o mouse para a posição desejada.")
        self.label.pack(pady=10)

        self.coord_label = tk.Label(root, text="X=0, Y=0", font=('Arial', 14))
        self.coord_label.pack(pady=10)

        self.quit_button = tk.Button(root, text="Sair", command=self.quit_app)
        self.quit_button.pack(pady=10)

        self.update_coordinates()

    def update_coordinates(self):
        try:
            x, y = pyautogui.position()
            self.coord_label.config(text=f"X={x}, Y={y}")
            self.root.after(100, self.update_coordinates)
        except tk.TclError:
            pass  # Captura o erro quando o programa é encerrado

    def quit_app(self):
        self.root.destroy()  # Encerra a aplicação de forma limpa


if __name__ == "__main__":
    root = tk.Tk()
    app = MousePositionApp(root)
    root.mainloop()
