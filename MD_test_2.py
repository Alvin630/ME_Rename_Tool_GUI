import tkinter as tk
from tkinter import filedialog, Toplevel, messagebox
import markdown
import os

class MarkdownViewerApp:
    def __init__(self, root, default_md_file):
        self.root = root
        self.root.title("Markdown Viewer")

        # # Main window button to manually open Markdown files
        # self.open_button = tk.Button(self.root, text="Open Another Markdown File", command=self.open_file)
        # self.open_button.pack(pady=10)

        # Automatically open the designated Markdown file in a Toplevel window
        if os.path.exists(default_md_file):
            self.show_markdown_in_new_window(default_md_file)
        else:
            messagebox.showerror("Error", f"Default Markdown file not found:\n{default_md_file}")

    # def open_file(self):
    #     # Allow user to open a different Markdown file
    #     file_path = filedialog.askopenfilename(
    #         title="Select a Markdown File",
    #         filetypes=[("Markdown Files", "*.md"), ("All Files", "*.*")]
    #     )
    #     if file_path:
    #         self.show_markdown_in_new_window(file_path)

    def show_markdown_in_new_window(self, file_path):
        # Read and convert the Markdown file to HTML
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                md_content = file.read()
            html_content = markdown.markdown(md_content)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read the file: {e}")
            return

        # Create a new Toplevel window
        new_window = Toplevel(self.root)
        new_window.title(f"Markdown Viewer - {os.path.basename(file_path)}")

        # Add a scrollable Text widget to display the content
        text_widget = tk.Text(new_window, wrap=tk.WORD)
        text_widget.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

        # Insert the HTML (rendered Markdown) into the Text widget
        text_widget.insert(tk.END, html_content)
        text_widget.configure(state="disabled")  # Make the content read-only


if __name__ == "__main__":
    # Path to the designated Markdown file to open on startup
    default_md_file = r"C:\Users\AlvinYT_Liang\Desktop\AutomaticTool\重新命名\ME rename tool\ME_Other_Folder_Structure.md"

    # Create the main Tkinter window
    root = tk.Tk()
    app = MarkdownViewerApp(root, default_md_file)
    root.mainloop()
