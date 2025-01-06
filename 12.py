import os
import win32com.client
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

class RemoveAltTextApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Anplify PowerPoint Alt Text Removal")
        self.configure(bg="#f0f0f0")

        # Set initial window size (adjust as needed)
        self.geometry("550x350")  
        self.resizable(False, False)

        self.font_family = "Arial"
        self.font_size_heading = 14
        self.font_size_text = 10

        self.presentation_path = None
        self.powerpoint = None

        self.create_widgets()

        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def create_widgets(self):
        # Main Frame
        self.main_frame = tk.Frame(self, bg="#f0f0f0")
        self.main_frame.pack(padx=20, pady=20, fill="both", expand=True)

        # Title Label
        self.label_title = tk.Label(
            self.main_frame,
            text="Anplify PowerPoint Alt Text Removal",
            font=(self.font_family, self.font_size_heading, "bold"),
            bg="#f0f0f0",
            fg="#333333"
        )
        self.label_title.pack(pady=(0, 15))

        # --- File Selection Section ---
        self.file_frame = tk.Frame(self.main_frame, bg="#f0f0f0")
        self.file_frame.pack(fill="x")

        # Browse Button
        self.button_browse = tk.Button(
            self.file_frame,
            text="Select PowerPoint File",
            command=self.select_file,
            font=(self.font_family, self.font_size_text),
            bg="#007bff",
            fg="white",
            relief="flat",
            padx=10,
            pady=5,
        )
        self.button_browse.pack(side="left")

        # Status Label
        self.label_status = tk.Label(
            self.file_frame,
            text="",
            font=(self.font_family, self.font_size_text, "italic"),
            bg="#f0f0f0",
            fg="#555555",
        )
        self.label_status.pack(side="left", padx=(10, 0))
        
        # --- Progress Bar ---
        self.progress_bar = ttk.Progressbar(
            self.main_frame,
            orient="horizontal",
            mode="indeterminate", 
            length=400,
        )
        self.progress_bar.pack(pady=(15,10))
        self.progress_bar.pack_forget() # Hide initially


        # --- Process Button ---
        self.button_process = tk.Button(
            self.main_frame,
            text="Remove Alt Text",
            command=self.process_file,
            state=tk.DISABLED,
            font=(self.font_family, self.font_size_text, "bold"),
            bg="#28a745",
            fg="white",
            relief="flat",
            padx=15,
            pady=8,
        )
        self.button_process.pack(pady=(10,15))

        # Output Label
        self.label_output = tk.Label(
            self.main_frame,
            text="",
            wraplength=450,
            font=(self.font_family, self.font_size_text),
            bg="#f0f0f0",
            fg="#333333",
        )
        self.label_output.pack()

    def select_file(self):
        self.presentation_path = filedialog.askopenfilename(
            filetypes=[("PowerPoint files", "*.pptx")]
        )
        if self.presentation_path:
            self.label_status.config(text=f"Selected: {os.path.basename(self.presentation_path)}")
            self.button_process.config(state=tk.NORMAL)
            self.label_output.config(text="") # Clear any previous output message

    def process_file(self):
        if not self.presentation_path:
            messagebox.showerror("Error", "No file selected!")
            return

        output_path = self.presentation_path.replace(".pptx", "_no_alt_text.pptx")

        try:
            # Show and start progress bar
            self.progress_bar.pack(pady=(15,10))
            self.progress_bar.start()

            # Remove alt text using win32com
            result = self.remove_alt_text(self.presentation_path, output_path)

            if result:
                self.label_output.config(
                    text=f"Alt text removed successfully!\nFile saved as: {os.path.basename(output_path)}",
                    fg="green"
                )
            else:
                messagebox.showerror("Error", "Alt text removal failed.")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

        finally:
            # Stop and hide progress bar
            self.progress_bar.stop()
            self.progress_bar.pack_forget()

    def remove_alt_text(self, input_path, output_path):
        try:
            # Open PowerPoint
            self.powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            # self.powerpoint.Visible = True  # Optional: Set to True to see PowerPoint
            presentation = self.powerpoint.Presentations.Open(input_path)

            # Iterate through slides and shapes
            for slide in presentation.Slides:
                for shape in slide.Shapes:
                    if hasattr(shape, "AlternativeText"):
                        shape.AlternativeText = ""
                    if hasattr(shape, "Title"):
                        shape.Title = ""
                    # Note: 'Description' attribute is not commonly used for alt text in PowerPoint.

            # Save the updated file
            presentation.SaveAs(output_path)
            presentation.Close()
            return True

        except Exception as e:
            print(f"Error during alt text removal: {e}")
            return False

        finally:
            if self.powerpoint:
                self.powerpoint.Quit()
                self.powerpoint = None  # Reset the variable

    def on_close(self):
        """Handles the window close event."""
        if self.powerpoint:
            self.powerpoint.Quit()
        self.destroy()

if __name__ == "__main__":
    app = RemoveAltTextApp()
    app.mainloop()