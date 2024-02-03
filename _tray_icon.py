import tkinter as tk
from PIL import Image
from pystray import Icon as TrayIcon, MenuItem as item
import threading

# Global variable for the tray icon
tray_icon = None


def create_window():
    window = tk.Tk()
    window.title("Tray Application")

    label = tk.Label(window, text="This is the main window.")
    label.pack()

    def hide_window():
        # Schedule the window to be hidden in the main thread
        window.after(0, window.withdraw)

    # Bind the minimize event
    window.bind(
        "<Unmap>", lambda event: hide_window() if window.state() == "iconic" else None
    )

    window.protocol("WM_DELETE_WINDOW", hide_window)

    return window


def show_window(icon, window):
    # Schedule the window to be de-iconified in the main thread
    window.after(0, lambda: window.deiconify())


def create_tray_icon(window, icon_image):
    global tray_icon

    # Define a function to quit the application.
    def quit_window(icon, window):
        icon.stop()
        window.after(0, window.destroy)

    # Define the menu for the system tray icon.
    menu = (
        item("Show", lambda: show_window(tray_icon, window)),
        item("Quit", lambda: quit_window(tray_icon, window)),
    )
    tray_icon = TrayIcon("Tray", icon_image, "My Tray App", menu)

    def run_icon():
        tray_icon.run()

    # Run the icon in a separate thread to avoid blocking the main thread.
    icon_thread = threading.Thread(target=run_icon)
    icon_thread.start()


if __name__ == "__main__":
    # Create the main window.
    window = create_window()

    # Load an image for the tray icon (this path should point to your icon image).
    icon_image = Image.open("your_icon.png")
    icon_image = icon_image.resize((16, 16), Image.Resampling.LANCZOS)

    # Create the tray icon.
    create_tray_icon(window, icon_image)

    window.mainloop()
