import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from PIL import Image, ImageTk
from tkinter.font import Font
import pprint
import platform
import subprocess
import shutil
from tkinterdnd2 import TkinterDnD, DND_FILES

class FileExplorerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("[ME] File Explorer & Rename Tool V-2.2.0")

        self.large_font = Font(family="Helvetica", size=12, weight="bold")

        # Frame for the first row of buttons
        self.frame1 = tk.Frame(self.root)
        self.frame1.pack(anchor="w",pady=5)

        # Frame for the second row of buttons
        self.frame2 = tk.Frame(self.root)
        self.frame2.pack(anchor="w",pady=5)

        self.create_folder = tk.Button(self.frame1, text="Create Folder",bg="#F9F900", command=self.trigger_folder_creation)
        self.create_folder.pack(side=tk.LEFT, padx=15)

        self.load_button = tk.Button(self.frame1, text="Load Directory", command=self.load_directory)
        self.load_button.pack(side=tk.LEFT, padx=15)

        self.expand_button = tk.Button(self.frame2, text="【+】", command=self.expand_all)
        self.expand_button.pack(side=tk.LEFT, padx=5)

        self.shrink_button = tk.Button(self.frame2, text="【-】", command=self.shrink_all)
        self.shrink_button.pack(side=tk.LEFT, padx=5)

        self.delete_button = tk.Button(self.frame2, text="Delete", command=self.delete_selected)
        self.delete_button.pack(side=tk.LEFT, padx=15)
        self.open_button = tk.Button(self.frame2, text="Open File", command=self.open_selected)
        self.open_button.pack(side=tk.LEFT, padx=15)
        
        self.rename_button = tk.Button(self.frame2, text="Rename Images in All Folders", command=self.rename_images)
        self.rename_button.pack(side=tk.LEFT, padx=15)

        self.path_label = tk.Label(self.frame1, text="", font=("Helvetica", 14))
        self.path_label.pack(side=tk.LEFT, padx=5)

        self.folder_tree_frame = tk.Frame(root)
        self.folder_tree_frame.pack(side=tk.LEFT, fill=tk.Y)
        self.folder_tree_scroll = tk.Scrollbar(self.folder_tree_frame)
        self.folder_tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.folder_tree = ttk.Treeview(self.folder_tree_frame, yscrollcommand=self.folder_tree_scroll.set, height=20, style="Folder.Treeview")
        self.folder_tree.pack(side=tk.LEFT, fill=tk.Y)
        self.folder_tree_scroll.config(command=self.folder_tree.yview)

        self.image_tree_frame = tk.Frame(root)
        self.image_tree_frame.pack(side=tk.RIGHT, expand=True, fill=tk.BOTH)
        self.image_tree_scroll = tk.Scrollbar(self.image_tree_frame)
        self.image_tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.image_tree = ttk.Treeview(self.image_tree_frame, yscrollcommand=self.image_tree_scroll.set, height=20, style="Image.Treeview")
        self.image_tree.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)
        self.image_tree_scroll.config(command=self.image_tree.yview)

        self.folder_tree.bind("<ButtonRelease-1>", self.on_folder_click)
        self.image_tree.bind("<Double-1>", self.on_double_click)

        # Enable drag-and-drop in the folder tree
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.on_drop)

        self.dir_path = ""
        self.image_cache = {}

        self.folder_tree.column("#0", width=550)
        self.image_tree.column("#0", width=450)

        style = ttk.Style()
        style.configure('Treeview', rowheight=20)
        style.configure('Folder.Treeview', rowheight=50)
        style.configure('Image.Treeview', rowheight=300)

        self.folder_tree.tag_configure('folder', font=self.large_font)
        self.image_tree.tag_configure('image', font=self.large_font)
        self.image_tree.tag_configure('file', font=self.large_font)
        # Assuming `self.folder_tree` is your Treeview
        self.folder_tree.tag_configure('folder', foreground='blue',font=('Helvetica', 12, 'bold'))  # Folder items will be yellow
        self.folder_tree.tag_configure('file', foreground='black',font=('Helvetica', 10, 'bold'))   # File items will be black

        self.default_naming_patterns = {
            "Basic Function Test": {"": ["Start", "Finish"]},
            "IEC 60068_Low Temperature Power ON_OFF Test": {"": ["Setup", "Start", "finish"]},
            "IEC 60068_High Temperature High Humidity Operation Test": {"": ["Setup", "Start", "finish"]},
            "IEC 60068_High Temperature Operation Test": {"": ["Setup", "Start", "finish"]},
            "IEC 60068_High Temperature Power ON_OFF Test": {"": ["Setup", "Start", "finish"]},
            "IEC 60068_Temperature Cycle Operation Test": {"": ["Setup", "Start", "finish"]},
            "EN 50155_Low Temperature Start-up Test": {"": ["Setup", "Start", "finish"]},
            "EN 50155_Dry Heat Thermal Test": {"": ["Setup", "Start", "finish"]},
            "EN 50155_Cyclic Damp Heat Test": {"": ["Setup", "Start", "finish"]},
            "IEC 60945_Cold Test": {"": ["Setup", "Start", "finish"]},
            "IEC 60945_Dry Heat Test": {"": ["Setup", "Start", "finish"]},
            "IEC 60945_Damp Heat Test": {"": ["Setup", "Start", "finish"]},
            "IEC 61850_Cold Test Minimum Storage Temperature": {"": ["Setup", "Start", "finish"]},
            "IEC 61850_Dry Heat Test Maximum Storage Temperature": {"": ["Setup", "Start", "finish"]},
            "IEC 61850_Cold Test Operational": {"": ["Setup", "Start", "finish"]},
            "IEC 61850_Dry Heat Test Operational": {"": ["Setup", "Start", "finish"]},
            "IEC 61850_Change of Temperature Test": {"": ["Setup", "Start", "finish"]},
            "IEC 61850_Damp Heat Cyclic Test": {"": ["Setup", "Start", "finish"]},
            "IEC 61850_Damp Heat Steady State": {"": ["Setup", "Start", "finish"]},
            "DNV_Dry Heat Test": {"": ["Setup", "Start", "finish"]},
            "DNV_Cold Test": {"": ["Setup", "Start", "finish"]},
            "DNV_Damp Heat Test": {"": ["Setup", "Start", "finish"]},
            "IEC 60945&DNV_Dry Heat Test": {"": ["Setup", "Start", "finish"]},
            "IEC 60945&DNV_Cold Test": {"": ["Setup", "Start", "finish"]},
            "IEC 60945&DNV_Damp Heat Test": {"": ["Setup", "Start", "finish"]}
        }

        self.burnin_patterns = {
            "IEC 60068_Low Temperature Power ON_OFF Test": ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "BurnIn_1hr"],
            "IEC 60068_High Temperature High Humidity Operation Test": ["BurnIn"],
            "IEC 60068_High Temperature Operation Test": ["BurnIn"],
            "IEC 60068_High Temperature Power ON_OFF Test": ["BurnIn","1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "BurnIn_1hr"],
            "IEC 60068_Temperature Cycle Operation Test": ["BurnIn"],
            "EN 50155_Low Temperature Start-up Test": ["PerfT_S_1", "PerfT_S_2","On_1", "On_2", "PerfT_E_1","PerfT_E_2"],
            "EN 50155_Dry Heat Thermal Test": ["PerfT_S_1","PerfT_S_2", "On_1", "On_2","On_3","On_4", "PerfT_E_1","PerfT_E_2"],
            "EN 50155_Cyclic Damp Heat Test": ["PerfT_S_1", "PerfT_S_2","On", "BurnIn","PerfT_E_1","PerfT_E_2"],
            "IEC 60945_Cold Test": ["BurnIn"],
            "IEC 60945_Dry Heat Test": ["BurnIn"],
            "IEC 60945_Damp Heat Test": ["BurnIn"],
            "IEC 61850_Cold Test Minimum Storage Temperature": ["BurnIn"],
            "IEC 61850_Dry Heat Test Maximum Storage Temperature": ["BurnIn"],
            "IEC 61850_Cold Test Operational": ["BurnIn"],
            "IEC 61850_Dry Heat Test Operational": ["BurnIn"],
            "IEC 61850_Change of Temperature Test": ["BurnIn"],
            "IEC 61850_Damp Heat Cyclic Test": ["BurnIn"],
            "IEC 61850_Damp Heat Steady State": ["BurnIn"],
            "DNV_Dry Heat Test": ["BurnIn"],
            "DNV_Cold Test": ["BurnIn"],
            "DNV_Damp Heat Test": ["FT_1", "FT_2", "FT_3"],
            "IEC 60945&DNV_Dry Heat Test": ["BurnIn"],
            "IEC 60945&DNV_Cold Test": ["BurnIn"],
            "IEC 60945&DNV_Damp Heat Test": ["FT_1", "FT_2", "FT_3"]
        }

#創建資料夾結構函數       
# Function to create the folder structure
    def create_folder_structure(self, test_case_name, test_sample_chamber_names, selected_executions, burn_in_samples, base_path):
        clean_test_case_name = test_case_name.strip()

        # Create the outer folder for the test case
        dir_path = os.path.join(base_path, clean_test_case_name)
        os.makedirs(dir_path, exist_ok=True)

        # Debugging: Print the test_sample_chamber_names
        print(f"Sample-Chamber pairs: {test_sample_chamber_names}")

        # If test_sample_chamber_names is empty, proceed without creating sample folders
        if not test_sample_chamber_names:
            # Iterate over each selected test execution
            for test_execution in selected_executions:
                test_execution_names = self.get_test_execution_names(test_execution)

                for exec_name in test_execution_names:
                    exec_dir = os.path.join(dir_path, exec_name)
                    os.makedirs(exec_dir, exist_ok=True)

                    # Create the "BurnIn" folder and subfolders for burn-in samples
                    burnin_dir = os.path.join(exec_dir, "BurnIn")
                    os.makedirs(burnin_dir, exist_ok=True)

                    if burn_in_samples and burn_in_samples > 1:
                        for i in range(1, burn_in_samples + 1):
                            burn_in_sample_dir = os.path.join(burnin_dir, str(i))
                            os.makedirs(burn_in_sample_dir, exist_ok=True)
        else:
            # Create folder structure for each sample and chamber combination
            for item in test_sample_chamber_names:
                if isinstance(item, tuple):
                    # If the item is a tuple (sample_name, chamber_name)
                    sample_name, chamber_name = item
                    # If there's no sample name, use only the chamber name for folder naming
                    if sample_name:
                        folder_name = f"{sample_name}_{chamber_name}"
                    else:
                        folder_name = chamber_name
                else:
                    # If the item is just a string (chamber_name only)
                    folder_name = item

                # Debugging: Print folder creation info
                print(f"Creating folder: {folder_name}")

                # Create the folder for this sample/chamber
                test_sample_dir = os.path.join(dir_path, folder_name)
                os.makedirs(test_sample_dir, exist_ok=True)

                # Iterate over each selected test execution
                for test_execution in selected_executions:
                    test_execution_names = self.get_test_execution_names(test_execution)

                    for exec_name in test_execution_names:
                        exec_dir = os.path.join(test_sample_dir, exec_name)
                        os.makedirs(exec_dir, exist_ok=True)

                        # Create the "BurnIn" folder and subfolders for burn-in samples
                        burnin_dir = os.path.join(exec_dir, "BurnIn")
                        os.makedirs(burnin_dir, exist_ok=True)

                        if burn_in_samples and burn_in_samples > 1:
                            for i in range(1, burn_in_samples + 1):
                                burn_in_sample_dir = os.path.join(burnin_dir, str(i))
                                os.makedirs(burn_in_sample_dir, exist_ok=True)

        # Notify the user of success and open the folder
        messagebox.showinfo("Success", f"Folder structure created at: {dir_path}")
        # Set the dir_path to the newly created directory
        self.dir_path = dir_path
        # Update the tree view to reflect the newly created folder structure
        self.refresh_views()
        # Select the created folder in the tree view
        self.reselect_folder(dir_path)
        # Open the folder in the file explorer (optional, if you want to automatically open it)
        self.open_folder(dir_path)

        # Update the path label
        self.path_label.config(text="路徑：" + dir_path)


#定義test execution name 函數
    def get_test_execution_names(self, test_execution):
        """Helper function to map test_execution to actual folder names."""
        if test_execution == "[Shacker] IEC 60068":
            return [
                "IEC 60068_Search Test(Sine Vib)",
                "IEC 60068_Endurance Test(Sine Vib)",
                "IEC 60068_Random Test(Rand Vib)",
                "IEC 60068_Shock Test(Half sine)",
            ]
        elif test_execution == "[Shacker] EN 50155":
            return [
                "EN 50155_LongLife Test(Rand Vib)",
                "EN 50155_Shock Test(Half sine)",
                "EN 50155_Functional Test(Rand Vib)"
            ]
        elif test_execution == "[Shacker] IEC 60945":
            return [
                "IEC 60945_Vibration Test(Sine Vib)",
            ]
        elif test_execution == "[Shacker] IEC 61850":
            return [
                "IEC 61850_Vibration Resonance Test(Sine Vib)",
                "IEC 61850_Vibration Endurance Test(Sine Vib)",
                "IEC 61850_Shock Response Test(Half sine)",
                "IEC 61850_Shock WithStand Test(Half sine)",
                "IEC 61850_Bump Test",
                "IEC 61850_Seismic Test"
            ]
        elif test_execution == "[Shacker] DNV":
            return [
                "DNV_Sweep Sine Test",
                "DNV_Wide Band Random Test",
            ]
        elif test_execution == "[Shacker] IEC 60945&DNV":
            return [
                "IEC 60945&DNV_Vibration Test(Sine Vib)",
            ]
        elif test_execution == "[ISTA] Shacker&Drop":
            return [
                "ISTA_Package Vib Test",
                "ISTA_Package Drop Test"
            ]
        elif test_execution == "[Shacker] IEC 60945&DNV":
            return [
                "IEC 60945&DNV_Vibration Test(Sine Vib) ",
            ]
        else:
            return [test_execution]

#創建資料夾函數
    def open_folder(self, path):
        # Detect the operating system and open the folder accordingly
        if platform.system() == "Windows":
            os.startfile(path)
        elif platform.system() == "Darwin":  # macOS
            subprocess.Popen(["open", path])
        elif platform.system() == "Linux":
            subprocess.Popen(["xdg-open", path])
        else:
            messagebox.showerror("Error", "Unsupported operating system.")

#定義使用者輸入並創建資料夾函數
    def get_user_input_and_create_folders(self):
        # Helper functions to get valid inputs
        def ask_valid_string(prompt_title, prompt_text):
            while True:
                user_input = simpledialog.askstring(prompt_title, prompt_text, parent=self.root)
                if user_input is None:
                    messagebox.showwarning("Input Cancelled", "使用者取消!  ")
                    return None
                elif user_input.strip() == "":
                    messagebox.showwarning("Invalid Input", "此為必填，請輸入有效字元!")
                else:
                    return user_input

        def ask_valid_integer(prompt_title, prompt_text):
            while True:
                user_input = simpledialog.askinteger(prompt_title, prompt_text, parent=self.root)
                if user_input is None:
                    messagebox.showwarning("Input Cancelled", "使用者取消!  ")
                    return None
                else:
                    return user_input

        def test_sample_chamber():
            sample_rows = []

            # Create the input window
            input_window = tk.Toplevel(self.root)
            input_window.title("Input Test Case Details")
            input_window.geometry("600x400")

            # Frame to hold sample rows
            sample_frame = tk.Frame(input_window)
            sample_frame.grid(row=0, column=0, columnspan=2, pady=10, padx=5)

            # Function to handle adding a new row
            def add_sample_row(frame, sample_rows):
                # Add a new sample row to the frame
                row_frame = tk.Frame(frame)
                row_frame.pack(pady=5)

                tk.Label(row_frame, text="Sample Name: ").pack(side=tk.LEFT, padx=5)
                sample_entry = tk.Entry(row_frame, width=20)
                sample_entry.pack(side=tk.LEFT, padx=5)

                tk.Label(row_frame, text="Mount Model: ").pack(side=tk.LEFT, padx=5)
                mount_var = tk.StringVar(value="Wall Mount")
                mount_menu = tk.OptionMenu(row_frame, mount_var, " ", "Chamber_A", "Chamber_B", "Chamber_C", "Chamber_D",
                                            "Chamber_E", "Chamber_F", "Chamber_G", "Chamber_H", "Chamber_I", "Chamber_J",
                                            "Chamber_K", "Chamber_L", "Chamber_M", "Chamber_N", "Chamber_O", "Chamber_P",
                                            "Chamber_Q", "Chamber_T")
                mount_menu.pack(side=tk.LEFT, padx=5)

                # Function to save the row data
                def save_row():
                    sample_name = sample_entry.get().strip()
                    chamber_name = mount_var.get()

                    if not sample_name and chamber_name.strip() != " ":  # If no sample name but chamber selected
                        if chamber_name not in [chamber for _, chamber in sample_rows]:
                            sample_rows.append((None, chamber_name))  # Add only the chamber model
                    elif sample_name and chamber_name.strip() != "" and (sample_name, chamber_name) not in sample_rows:
                        sample_rows.append((sample_name, chamber_name))  # Add both sample name and chamber model

                # Bind the row saving to the <Return> key
                sample_entry.bind("<Return>", lambda event: save_row())

                # Bind the row saving to focus out (optional)
                sample_entry.bind("<FocusOut>", lambda event: save_row())

            # Add the first sample row when the window opens
            add_sample_row(sample_frame, sample_rows)

            # Button to add more sample rows
            tk.Button(input_window, text="Add(+) ", command=lambda: add_sample_row(sample_frame, sample_rows)).grid(
                row=1, column=0, columnspan=2, pady=5)

            # Function to handle the window closure behavior (close button X)
            def on_close():
                """Handle when the user closes the window via the 'X' button."""
                messagebox.showwarning("Input Cancelled", "使用者取消!")
                input_window.destroy()  # Close the window
                return None  # Ensure it returns None when the window is closed

            # Adding the close button functionality (X)
            input_window.protocol("WM_DELETE_WINDOW", on_close)

            # Adding the skip button functionality
            def skip_and_close():
                input_window.destroy()  # Close the window and return to the next step
                return []
            tk.Button(input_window, text="Skip", command=skip_and_close).grid(row=3, column=1, pady=10)

            # Button to confirm and close the window
            def confirm_and_close():
                # Save the last row before closing (if not already saved)
                for child in sample_frame.winfo_children():
                    sample_entry_widget = child.winfo_children()[1]  # Assuming Entry is at index 1
                    chamber_var_widget = child.winfo_children()[3]  # Assuming OptionMenu is at index 3

                    sample_name = sample_entry_widget.get().strip()
                    chamber_name = chamber_var_widget.cget("text")

                    # Check for valid input
                    if not sample_name and chamber_name.strip() != " " and (None, chamber_name) not in sample_rows:
                        sample_rows.append((None, chamber_name))  # Add only the chamber name if no sample name
                    elif sample_name and chamber_name.strip() != "" and (sample_name, chamber_name) not in sample_rows:
                        sample_rows.append((sample_name, chamber_name))  # Add both sample name and chamber

                input_window.destroy()  # Close the window

            tk.Button(input_window, text="Enter", command=confirm_and_close).grid(row=2, column=1, pady=10)

            # Wait until the window is closed
            input_window.wait_window()

            # Process sample_rows to remove None for folders
            processed_sample_rows = [chamber if sample is None else f"{sample}_{chamber}" for sample, chamber in sample_rows]
            return processed_sample_rows if processed_sample_rows else []

        def select_test_executions():
            selected_executions = []
            options = ["[Shacker] IEC 60068", "[Shacker] EN 50155", "[Shacker] IEC 60945", "[Shacker] IEC 61850", "[Shacker] DNV",
                        "IEC 60945&DNV", "[Shacker] IEC 60945&DNV","[ISTA] Shacker&Drop" ]
            #Function to handle the window closure behavior (close button X)

            
            # Create a new window for the multi-select dropdown menu
            execution_window = tk.Toplevel(self.root)
            execution_window.title("Test Executions")
            execution_window.attributes("-topmost", True)

            tk.Label(execution_window, text="選擇欲測試之EE測項:").pack(pady=10)

            # Multi-select dropdown menu for test execution
            test_executions = tk.Listbox(execution_window, selectmode=tk.MULTIPLE, height=8, width=40)
            for item in options:
                test_executions.insert(tk.END, item)
            test_executions.pack(pady=10)

            def on_close():
                """Handle when the user closes the window via the 'X' button."""
                messagebox.showwarning("Input Cancelled", "使用者取消!")
                execution_window.destroy()  # Close the window
                return None  # Ensure it returns None when the window is closed

            # Adding the close button functionality (X)
            execution_window.protocol("WM_DELETE_WINDOW", on_close)

            def confirm_executions():
                selected_indices = test_executions.curselection()
                for i in selected_indices:
                    selected_executions.append(options[i])  # Get the actual string values
                if not selected_executions:
                    messagebox.showwarning("Invalid Input", "此為必選項目,請選擇測項!")
                    
                else:
                    execution_window.destroy()

            tk.Button(execution_window, text="Enter", command=confirm_executions).pack(pady=10)

            # Wait until the window is closed
            execution_window.wait_window()

            return selected_executions if selected_executions else None

        # Get test case name
        test_case_name = ask_valid_string("Test Case Name", "請輸入測試專案號碼+測試樣品名(Jira)!     ").strip()
        if test_case_name is None:
            return

        # Get test sample and chamber model names (allow skipping)
        test_sample_chamber_names = test_sample_chamber()

        # If the user skipped the sample input, continue with an empty list
        if test_sample_chamber_names is None:
            return

        # Get test executions
        selected_executions = select_test_executions()
        if selected_executions is None:
            return

        # Get burn-in samples
        burn_in_samples = ask_valid_integer("BurnIn Number", "請輸入燒機數(燈號辨認請輸入'1',外接螢幕視需求輸入數量!)")
        if burn_in_samples is None:
            return

        # Ask the user to choose a directory where the structure will be created
        base_path = filedialog.askdirectory(title="Select Base Directory")
        if not base_path:
            messagebox.showerror("Error", "請選擇欲創建資料夾位置!      ")
            return

        # Create the folder structure
        self.create_folder_structure(test_case_name, test_sample_chamber_names, selected_executions, burn_in_samples, base_path)


        
    #定義TK GUI 創建資料夾函數按鈕功能函數
    def trigger_folder_creation(self):
        # Trigger the folder creation process
        self.get_user_input_and_create_folders()


    def load_directory(self):
        self.dir_path = filedialog.askdirectory()
        if self.dir_path:
            self.path_label.config(text="路徑：" + self.dir_path)
            self.folder_tree.delete(*self.folder_tree.get_children())
            self.image_tree.delete(*self.image_tree.get_children())
            self.load_folders(self.dir_path)

    #新增左邊資料夾區之內部檔案表現
    def load_folders(self, path, parent=""):
        # First, store image files and other files separately
        image_files = []
        other_files = []
        folders = []
        
        for item in os.listdir(path):
            abs_path = os.path.join(path, item)

            if os.path.isdir(abs_path):
                # Store the folder for later insertion
                folders.append(item)
            else:
                # If the item is a file, categorize it
                file_extension = os.path.splitext(item)[1].lower()
                if file_extension in ['.png', '.jpg', '.jpeg']:
                    image_files.append(item)
                else:
                    other_files.append(item)

        # Now insert image files first
        for image in image_files:
            self.folder_tree.insert(parent, 'end', text=image, tags=('file',))
        
        # Then insert the other files
        for file in other_files:
            self.folder_tree.insert(parent, 'end', text=file, tags=('file',))
        
        # Finally, insert folders and load their contents recursively
        for folder in folders:
            node = self.folder_tree.insert(parent, 'end', text=folder, open=False, tags=('folder',))
            self.load_folders(os.path.join(path, folder), node)

    def on_folder_click(self, event):
        # Get the current selection from the folder tree
        selection = self.folder_tree.selection()

        # Ensure that the selection is not empty to avoid IndexError
        if selection:
            selected_item = selection[0]
            folder_path = self.get_item_path(selected_item)

            # Check if the selected item is a folder
            if os.path.isdir(folder_path):
                self.load_files_and_images(folder_path)
            else:
                # Deselect the file item to prevent any action
                self.folder_tree.selection_remove(selected_item)
        else:
            print("No folder selected.")

    def load_files_and_images(self, path):
        self.image_tree.delete(*self.image_tree.get_children())
        for item in os.listdir(path):
            abs_path = os.path.join(path, item)
            if os.path.isfile(abs_path):
                file_extension = os.path.splitext(item)[1].lower()
                if file_extension in ['.png', '.jpg', '.jpeg']:
                    if abs_path not in self.image_cache:
                        img = Image.open(abs_path)
                        img.thumbnail((370, 370))
                        self.image_cache[abs_path] = ImageTk.PhotoImage(img)
                    image = self.image_cache[abs_path]
                    self.image_tree.insert('', 'end', text=item, image=image, tags=('image',))
                else:
                    self.image_tree.insert('', 'end', text=item, tags=('file',))
    #用在建構Folder tree
    def get_item_path(self, item_id):
        path_components = []
        while item_id:
            item_text = self.folder_tree.item(item_id, "text")
            path_components.append(item_text)
            item_id = self.folder_tree.parent(item_id)
        path_components.reverse()
        return os.path.join(self.dir_path, *path_components)
    
    #更新樹狀圖介面(待修正錯誤問題)
    def refresh_views(self):
        # Reselect the previously selected folder with a slight delay
        def refresh_with_delay():
            # Remember the currently selected folder
            selected_folder_id = self.folder_tree.selection()
            if selected_folder_id:
                selected_folder_path = self.get_item_path(selected_folder_id[0])
            else:
                selected_folder_path = self.dir_path

            # Clear and reload the folder and image trees
            self.folder_tree.delete(*self.folder_tree.get_children())
            self.image_tree.delete(*self.image_tree.get_children())
            self.load_folders(self.dir_path)

            # Expand all the folders first
            self.expand_all()

            # Reselect the previously selected folder
            self.reselect_folder(selected_folder_path)

            # Reload the images in the reselected folder
            self.load_files_and_images(selected_folder_path)

        # Use self.after to introduce a small delay (e.g., 100 ms)
        self.root.after(100, refresh_with_delay)

    def reselect_folder(self, path):
        # Recursively search for the folder by path and select it
        def recursive_select(node, target_path):
            item_path = self.get_item_path(node)
            if item_path == target_path:
                self.folder_tree.selection_set(node)
                self.folder_tree.see(node)
                return True
            for child in self.folder_tree.get_children(node):
                if recursive_select(child, target_path):
                    return True
            return False

        for child in self.folder_tree.get_children(''):
            if recursive_select(child, path):
                break
    #雙擊重新命名功能
    def on_double_click(self, event):
        item_id = self.image_tree.selection()[0]
        item_text = self.image_tree.item(item_id, "text")
        item_path = self.get_image_item_path(item_id)
        if not item_path:
            return  # Stop if the path could not be retrieved

        original_extension = os.path.splitext(item_text)[1]

        new_name = CustomSimpledialog(self.root, title="Rename", message="請輸入新名稱：", initialvalue=item_text).result
        if new_name:
            new_extension = os.path.splitext(new_name)[1]
            if not new_extension:
                new_name += original_extension

            new_path = os.path.join(os.path.dirname(item_path), new_name)
            os.rename(item_path, new_path)
            # Refresh folder and image views
            self.refresh_views()

    #用在建構image tree & file tree
    def get_image_item_path(self, item_id):
        # Check if any folder is selected in the folder_tree
        folder_selection = self.folder_tree.selection()
        if not folder_selection:
            # No folder is selected, handle it gracefully (you can show a message if needed)
            messagebox.showerror("Error", "No folder selected!(請選擇資料夾路徑!)")
            return None
        # If folder is selected, construct the path
        selected_folder = self.get_item_path(self.folder_tree.selection()[0])
        image_name = self.image_tree.item(item_id, "text")
        return os.path.join(selected_folder, image_name)
    
    #用於拖曳檔案(複製)至GUI介面中    
    def on_drop(self, event):
        # Get the destination folder from the treeview selection
        dest_item = self.folder_tree.selection()
        dest_path = self.get_item_path(dest_item[0]) if dest_item else self.dir_path

        if os.path.isdir(dest_path):
            dropped_files = self.root.tk.splitlist(event.data)
            for file in dropped_files:
                try:
                    # Handle both file and folder copies
                    if os.path.isfile(file):
                        shutil.copy(file, dest_path)  # Copy file
                    elif os.path.isdir(file):
                        shutil.copytree(file, os.path.join(dest_path, os.path.basename(file)))  # Copy folder
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to copy file: {e}")
            self.load_files_and_images(dest_path)  # Refresh the files and images view
            self.refresh_views()

    # 刪除資料夾或檔案功能函數
    def delete_selected(self):
        try:
            if self.image_tree.selection():  # Check if an image or file is selected
                selected_item = self.image_tree.selection()[0]
                file_path = self.get_image_item_path(selected_item)
                if messagebox.askyesno("Delete", "Are you sure you want to delete this file?"):
                    os.remove(file_path)
                    self.image_tree.delete(selected_item)  # Remove the item from the tree view
                    self.refresh_views()
            elif self.folder_tree.selection():  # Check if a folder is selected
                selected_item = self.folder_tree.selection()[0]
                folder_path = self.get_item_path(selected_item)
                if messagebox.askyesno("Delete", "Are you sure you want to delete this folder and its contents?"):
                    shutil.rmtree(folder_path)  # Delete the folder and its contents
                    self.folder_tree.delete(selected_item)  # Remove the item from the tree view
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete: {e}")

    #展開與摺疊
    def expand_all(self):
        for item in self.folder_tree.get_children():
            self.folder_tree.item(item, open=True)
            self.expand_all_children(item)

    def expand_all_children(self, item):
        for child in self.folder_tree.get_children(item):
            self.folder_tree.item(child, open=True)
            self.expand_all_children(child)

    def shrink_all(self):
        for item in self.folder_tree.get_children():
            self.folder_tree.item(item, open=False)
            self.shrink_all_children(item)

    def shrink_all_children(self, item):
        for child in self.folder_tree.get_children(item):
            self.folder_tree.item(child, open=False)
            self.shrink_all_children(child)

# open file function 
    def open_selected(self):
        selected_item = self.image_tree.selection()
        if selected_item:
            item_path = self.get_image_item_path(selected_item[0])
            if os.path.isfile(item_path):
                try:
                    os.startfile(item_path)
                except OSError as e:
                    messagebox.showerror("錯誤", f"無法開啟圖片 {item_path}: {e}")
            elif os.path.isdir(item_path):
                try:
                    os.startfile(item_path)
                except OSError as e:
                    messagebox.showerror("錯誤", f"無法開啟資料夾 {item_path}: {e}")
   
#修正只搜尋最外層資料夾符合字典的檔案修改，無外部資料夾不進行修改的錯誤20241017 (修正中)
#修正對字典內對應名稱重複命名造成的排序問題(移除create_unique_name)
    def check_and_rename_images_in_folder(self, test_type, folder_path, naming_patterns):
        """
        Helper function to check and rename images in a folder.
        """
        files = sorted(os.listdir(folder_path))
        images = [f for f in files if f.lower().endswith(('.png', '.jpg', '.jpeg'))]

        if len(images) != len(naming_patterns):
            messagebox.showwarning("Mismatch Warning", f"路徑：'{folder_path}'\n影像數目不對！")
            return 0

        used_names = set()
        renamed_count = 0

        for image_name in images:
            file_name, file_extension = os.path.splitext(image_name)

            # Skip renaming if already matches the naming pattern
            if file_name in naming_patterns:
                used_names.add(file_name)
                continue

            # Find an unused valid name for renaming
            new_base_name = next((name for name in naming_patterns if name not in used_names), None)
            if new_base_name is None:
                continue

            new_name = new_base_name + ".jpg"
            src = os.path.join(folder_path, image_name)
            dst = os.path.join(folder_path, new_name)

            if os.path.exists(dst):
                # Handle potential name conflicts
                dst = self.create_unique_name(folder_path, new_base_name, file_extension)

            os.rename(src, dst)
            used_names.add(new_base_name)
            renamed_count += 1

        return renamed_count

    def check_and_rename_outer_layer(self):
        total_renamed = 0

        # Traverse directories in the main path
        outer_layer_folders = [d for d in os.listdir(self.dir_path) if os.path.isdir(os.path.join(self.dir_path, d))]

        for folder in outer_layer_folders:
            folder_path = os.path.join(self.dir_path, folder)

            # Check if it's the test_type folder (in case there's no outer layer)
            if folder in self.default_naming_patterns:
                # This folder might be a direct test_type folder, process it
                total_renamed += self.process_test_type_folder(self.dir_path, folder)

            else:
                # Treat it as an outer layer folder and look inside it for test_type folders
                for test_type, subfolders in self.default_naming_patterns.items():
                    test_type_path = os.path.join(folder_path, test_type)

                    if os.path.isdir(test_type_path):
                        # Rename images in the test_type folder
                        total_renamed += self.process_test_type_folder(folder_path, test_type)

        return total_renamed

    def process_test_type_folder(self, base_path, test_type):
        """
        Process images inside a test_type folder by renaming them according to the naming patterns.
        """
        total_renamed = 0
        subfolders = self.default_naming_patterns.get(test_type, {})

        for subfolder, naming_patterns in subfolders.items():
            folder_path = os.path.join(base_path, test_type, subfolder)
            if os.path.isdir(folder_path):
                total_renamed += self.check_and_rename_images_in_folder(test_type, folder_path, naming_patterns)

        return total_renamed

    def check_and_rename_burnin_folders(self, burnin_subfolder_count):
        total_renamed = 0
        processed_folders = set()  # Track folders that have been processed

        # Traverse directories
        outer_layer_folders = [d for d in os.listdir(self.dir_path) if os.path.isdir(os.path.join(self.dir_path, d))]

        for folder in outer_layer_folders:
            folder_path = os.path.join(self.dir_path, folder)

            # Check test types in both outer layer or direct path
            for test_type, subfolders in self.default_naming_patterns.items():
                for subfolder, _ in subfolders.items():
                    # Case 1: BurnIn folder is inside an outer layer folder
                    burnin_folder_path = os.path.join(folder_path, test_type, subfolder, "BurnIn")
                    
                    if not os.path.isdir(burnin_folder_path):
                        # Case 2: BurnIn folder is directly in the test_type folder without an outer layer
                        burnin_folder_path = os.path.join(self.dir_path, test_type, subfolder, "BurnIn")

                    if not os.path.isdir(burnin_folder_path):
                        continue

                    # Check if this BurnIn folder has already been processed
                    if burnin_folder_path in processed_folders:
                        continue  # Skip already processed folder

                    # Mark this folder as processed
                    processed_folders.add(burnin_folder_path)

                    # Process BurnIn images
                    if burnin_subfolder_count == 1:
                        # Handle BurnIn images directly within the main BurnIn folder
                        total_renamed += self.check_and_rename_images_in_folder(
                            test_type, burnin_folder_path, self.burnin_patterns.get(test_type, [])
                        )
                    else:
                        # Multiple BurnIn subfolders
                        total_renamed += self.rename_burnin_subfolders(
                            burnin_folder_path, burnin_subfolder_count, test_type
                        )

        return total_renamed

    def rename_burnin_subfolders(self, burnin_folder_path, burnin_subfolder_count, test_type):
        """
        Rename images in multiple BurnIn subfolders.
        """
        total_renamed = 0
        subfolders = [f for f in os.listdir(burnin_folder_path) if os.path.isdir(os.path.join(burnin_folder_path, f)) and f.isdigit()]
        subfolders = sorted(subfolders, key=int)[:burnin_subfolder_count]

        for subfolder in subfolders:
            subfolder_path = os.path.join(burnin_folder_path, subfolder)
            if os.path.isdir(subfolder_path):
                total_renamed += self.check_and_rename_images_in_folder(
                    test_type, subfolder_path, self.burnin_patterns.get(test_type, [])
                )

        return total_renamed

    def rename_images(self):
        total_renamed_outer = self.check_and_rename_outer_layer()
        burnin_subfolder_count = self.prompt_burnin_subfolder_count()
        total_renamed_burnin = self.check_and_rename_burnin_folders(burnin_subfolder_count)
        self.refresh_views()
        messagebox.showinfo("Rename Complete", f"Renamed {total_renamed_outer + total_renamed_burnin} images.")

    # Mock-up function for create_unique_name (you need to implement this)
    def create_unique_name(self, folder_path, base_name, extension):
        """
        This function will create a unique name if a file already exists.
        """
        counter = 1
        new_name = f"{base_name}_{counter}{extension}"
        while os.path.exists(os.path.join(folder_path, new_name)):
            counter += 1
            new_name = f"{base_name}_{counter}{extension}"
        return new_name

    def prompt_burnin_subfolder_count(self):
        while True:
            user_input = CustomSimpledialog(self.root, title="BurnIn Record Number", message="請輸入燒機資料夾數量 \n(按照建立資料夾時的數目輸入)：", initialvalue="").result
            if user_input is None:
                return 0
            try:
                return int(user_input)
            except ValueError:
                messagebox.showerror("Error!!", "請輸入整數！")

class CustomSimpledialog(simpledialog.Dialog):
    def __init__(self, parent, title, message, initialvalue=""):
        self.message = message
        self.initialvalue = initialvalue
        super().__init__(parent, title)

    def body(self, master):
        tk.Label(master, text=self.message).pack()
        self.entry = tk.Entry(master,width=50)
        self.entry.pack(padx=40,pady=10)
        self.entry.insert(0, self.initialvalue)
        return self.entry

    def apply(self):
        self.result = self.entry.get()

if __name__ == "__main__":
    root = TkinterDnD.Tk()  # Create the root window with drag-and-drop enabled
    app = FileExplorerApp(root)  # Pass the root to FileExplorerApp
    root.mainloop()  # Start the main event loop