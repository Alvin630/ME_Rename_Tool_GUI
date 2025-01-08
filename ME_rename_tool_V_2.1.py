import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from PIL import Image, ImageTk
from tkinter.font import Font
import pprint
import platform
import subprocess

class FileExplorerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("[ME] File Explorer & Rename Tool V-2.1.0")

        self.large_font = Font(family="Helvetica", size=14, weight="bold")

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

        self.arrange_folder=tk.Button(self.frame1, text="Folder arrange", command=self.browse_check_file)
        self.arrange_folder.pack(side=tk.LEFT, padx=15)

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
        
        self.dir_path = ""
        self.image_cache = {}

        self.folder_tree.column("#0", width=650)
        self.image_tree.column("#0", width=350)

        style = ttk.Style()
        style.configure('Treeview', rowheight=20)
        style.configure('Folder.Treeview', rowheight=50)
        style.configure('Image.Treeview', rowheight=325)

        self.folder_tree.tag_configure('folder', font=self.large_font)
        self.image_tree.tag_configure('image', font=self.large_font)
        self.image_tree.tag_configure('file', font=self.large_font)

        self.default_naming_patterns = {
            "Basic Function test": {"": ["Start", "Finish"]},
            "IEC 60068_Search(Sine Vib)": {"": ["Setup", "Start", "finish"]},
            "IEC 60068_Endurance(Sine Vib)": {"": ["Setup", "Start", "finish"]},
            "IEC 60068_Random test(Rand Vib)": {"": ["Setup", "Start", "finish"]},
            "IEC 60068_Shock Test(Half sine)": {"": ["Setup", "Start", "finish"]},
            "EN 50155_LongLife Test(Rand Vib)": {"": ["Setup", "Start", "finish"]},
            "EN 50155_Shock Testing(Half sine)": {"": ["Setup", "Start", "finish"]},
            "EN 50155_Functional Test(Rand Vib)": {"": ["Setup", "Start", "finish"]},
            "IEC 60945_Vibration(Sine Vib)": {"": ["Setup", "Start", "finish"]},
            "IEC 61850_Vibration resonance Test(Sine Vib)": {"": ["Setup", "Start", "finish"]},
            "IEC 61850_Vibration endurance Test_(Sine Vib)": {"": ["Setup", "Start", "finish"]},
            "IEC 61850_Shock response Test(Half sine)": {"": ["Setup", "Start", "finish"]},
            "IEC 61850_Shock WithStand Test(Half sine)": {"": ["Setup", "Start", "finish"]},
            "IEC 61850_Bump test": {"": ["Setup", "Start", "finish"]},
            "IEC 61850_Seismic test": {"": ["Setup", "Start", "finish"]},
            "DNV_Sweep Sine Test": {"": ["Setup", "Start", "finish"]},
            "DNV_Wide Band Random": {"": ["Setup", "Start", "finish"]},
            "IEC 60945&DNV_Vibration(Sine Vib)": {"": ["Setup", "Start", "finish"]},
        }

        self.burnin_patterns = {
            "IEC 60068_Low Temperature Power ON_OFF Test": ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "BurnIn_1hr"],
            "IEC 60068_High Temperature High Humidity Operation Test": ["BurnIn"],
            "IEC 60068_High Temperature Operation Test": ["BurnIn"],
            "IEC 60068_High Temperature Power ON_OFF Test": ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "BurnIn_1hr"],
            "IEC 60068_Temperature Cycle Operation Test": ["BurnIn"],
            "EN 50155_Low Temperature Start-up Test": ["PerfT_S", "On_1", "On_2", "Perf_E"],
            "EN 50155_Dry Heat Thermal Test": ["PerfT_S", "On_1", "On_2", "Perf_E"],
            "EN 50155_Cyclic Damp Heat Test": ["PerfT_S", "On_1", "On_2", "Perf_E"],
            "IEC 60945_Cold Test": ["BurnIn"],
            "IEC 60945_Dry heat Test": ["BurnIn"],
            "IEC 60945_Damp heat Test": ["BurnIn"],
            "IEC 61850_Cold Test Minimum Storage Temperature": ["BurnIn"],
            "IEC 61850_Dry Heat Test Maximum Storage Temperature": ["BurnIn"],
            "IEC 61850_Cold Test Operational": ["BurnIn"],
            "IEC 61850_Dry Heat Test Operational": ["BurnIn"],
            "IEC 61850_Change of Temperature Test": ["BurnIn"],
            "IEC 61850_Damp Heat Cyclic Test": ["BurnIn"],
            "IEC 61850_Damp Heat Steady State": ["BurnIn"],
            "DNV_Dry heat Test": ["BurnIn"],
            "DNV_Cold Test ": ["BurnIn"],
            "DNV_Damp heat Test": ["FT_1", "FT_2", "FT_3"],
            "IEC 60945&DNV_Dry heat Test": ["BurnIn"],
            "IEC 60945&DNV_Cold Test ": ["BurnIn"],
            "IEC 60945&DNV_Damp heat Test": ["FT_1", "FT_2", "FT_3"]
        }
    def create_folder_structure(self, test_case_name, test_sample_names, test_executions, burn_in_samples, subfolder_names, base_path):
        # Create the base directory for the test case
        test_case_dir = os.path.join(base_path, test_case_name)
        os.makedirs(test_case_dir, exist_ok=True)

        # Split the test_sample_names by comma
        sample_names_list = [name.strip() for name in test_sample_names.split(',')] if test_sample_names else ['']

        # Iterate over each sample name
        for test_sample_name in sample_names_list:
            # Set the base directory for the sample name (skip if empty)
            if test_sample_name:
                test_sample_dir = os.path.join(test_case_dir, test_sample_name)
            else:
                test_sample_dir = test_case_dir

            # Iterate over each selected test execution
            for test_execution in test_executions:
                # Mapping test_execution to actual folder names
                test_execution_names = []
                if test_execution == "Shacker IEC 60068":
                    test_execution_names = [
                        "IEC 60068_Search(Sine Vib)",
                        "IEC 60068_Endurance(Sine Vib)",
                        "IEC 60068_Random test(Rand Vib)",
                        "IEC 60068_Shock Test(Half sine)",
                    ]
                elif test_execution == "Shacker EN 50155":
                    test_execution_names = [
                        "EN 50155_LongLife Test(Rand Vib)",
                        "EN 50155_Shock Testing(Half sine)",
                        "EN 50155_Functional Test(Rand Vib)"
                    ]
                elif test_execution == "Shacker IEC 60945":
                    test_execution_names = [
                        "IEC 60945_Cold Test",
                        "IEC 60945_Dry heat Test",
                        "IEC 60945_Damp heat Test"
                    ]
                elif test_execution == "Shacker IEC 61850":
                    test_execution_names = [
                        "IEC 61850_Vibration resonance Test(Sine Vib)",
                        "IEC 61850_Vibration endurance Test_(Sine Vib)",
                        "IEC 61850_Shock response Test(Half sine)",
                        "IEC 61850_Shock WithStand Test(Half sine)",
                        "IEC 61850_Bump test",
                        "IEC 61850_Seismic test"
                    ]
                elif test_execution == "Shacker DNV":
                    test_execution_names = [
                        "DNV_Sweep Sine Test",
                        "DNV_Wide Band Random",
                    ]
                elif test_execution == "Shacker IEC 60945&DNV":
                    test_execution_names = [
                        "IEC 60945&DNV_Vibration(Sine Vib)",
                    ]
                elif test_execution == "Shacker&Drop ISTA 1A":
                    test_execution_names = [
                        "ISTA_Package Vib Test",
                        "ISTA_Package Drop Test",
                    ]
                elif test_execution == "EN 50155 Salt mist ":
                    test_execution_names = [
                        "Salt mist test",
                    ]
                elif test_execution == "Manual Test":
                    test_execution_names = [
                        "Interface connector durability",
                    ]
                elif test_execution == "IEC 60529 IP Test":
                    test_execution_names = [
                        "Ingress Protection 2X",
                        "Ingress Protection 3X",
                        "Ingress Protection 4X",
                        "Ingress Protection 5X",
                        "Ingress Protection 6X",
                        "Ingress Protection X3",
                        "Ingress Protection X4",
                        "Ingress Protection X5",
                        "Ingress Protection X6",
                        "Ingress Protection X7",
                        "Ingress Protection X8",
                        "IP老化測試(Temperature Cycle Test)"
                    ]
                elif test_execution == "IEC 60068 Storage Test":
                    test_execution_names = [
                        "High Temperature High Humidity Storage Test",
                        "Low Temperature Storage Test"
                    ]
                else:
                    test_execution_names = [test_execution]

                for exec_name in test_execution_names:
                    exec_dir = os.path.join(test_sample_dir, exec_name)
                    os.makedirs(exec_dir, exist_ok=True)

                    # Create subfolders inside the test execution directory
                    if subfolder_names:
                        for subfolder in subfolder_names:
                            subfolder_dir = os.path.join(exec_dir, subfolder.strip())
                            os.makedirs(subfolder_dir, exist_ok=True)

                    # Create the directories for each burn-in sample if more than 1 sample
                    if burn_in_samples > 1:
                        for i in range(1, burn_in_samples + 1):
                            burn_in_dir = os.path.join(exec_dir, str(i))
                            os.makedirs(burn_in_dir, exist_ok=True)

        messagebox.showinfo("Success", f"Folder structure created at: {base_path}")

        # Automatically open the folder after creation
        self.open_folder(test_case_dir)

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

    def get_user_input_and_create_folders(self):
        # Ask the user for inputs
        test_case_name = simpledialog.askstring("Input", "What's your test case name?", parent=self.root)
        test_sample_names = simpledialog.askstring(
            "Input", "What's your test sample name? (Use ',' to separate multiple names)", parent=self.root
        )

        # Create a new window for the multi-select dropdown menu
        execution_window = tk.Toplevel(self.root)
        execution_window.title("Select Test Executions")

        tk.Label(execution_window, text="Select Test Executions:").pack(pady=10)

        # Multi-select dropdown menu for test execution
        test_executions = tk.Listbox(execution_window, selectmode=tk.MULTIPLE, height=6)
        for item in ["Shacker IEC 60068", "Shacker EN 50155", "Shacker IEC 60945", "Shacker IEC 61850", "Shacker DNV", "Shacker IEC 60945&DNV",
                     "Shacker&Drop ISTA 1A","Salt mist EN 50155",]:
            test_executions.insert(tk.END, item)
        test_executions.pack(pady=10)

        selected_executions = []

        def confirm_executions():
            selected_indices = test_executions.curselection()
            for i in selected_indices:
                selected_executions.append(test_executions.get(i))
            execution_window.destroy()

        tk.Button(execution_window, text="Confirm", command=confirm_executions).pack(pady=10)

        # Wait until the window is closed
        execution_window.wait_window()

        burn_in_samples = simpledialog.askinteger("Input", "How many samples do you burn in? (e.g., 3)")

        # Create a window for selecting predefined subfolder names and allowing custom input
        subfolder_window = tk.Toplevel(self.root)
        subfolder_window.title("Select or Enter Subfolder Names")

        tk.Label(subfolder_window, text="Select Subfolders:").pack(pady=10)

        # Define some predefined subfolder options
        predefined_subfolders = ["X","Y","Z","測試前六面", "測試後六面", "外箱六面"]
        subfolder_vars = []

        # Create checkboxes for the predefined subfolder options
        for subfolder in predefined_subfolders:
            var = tk.BooleanVar()
            cb = tk.Checkbutton(subfolder_window, text=subfolder, variable=var)
            cb.pack(anchor="w")
            subfolder_vars.append((subfolder, var))

        # Add a text entry for custom subfolder names
        custom_subfolder_label = tk.Label(subfolder_window, text="Enter Custom Subfolders (',' separated):")
        custom_subfolder_label.pack(pady=10)
        custom_subfolder_entry = tk.Entry(subfolder_window)
        custom_subfolder_entry.pack(pady=10)

        selected_subfolders = []

        def confirm_subfolders():
            # Get the selected predefined subfolders
            for subfolder, var in subfolder_vars:
                if var.get():
                    selected_subfolders.append(subfolder)

            # Get custom subfolder names entered by the user
            custom_subfolders = custom_subfolder_entry.get()
            if custom_subfolders:
                selected_subfolders.extend([s.strip() for s in custom_subfolders.split(',')])

            subfolder_window.destroy()

        tk.Button(subfolder_window, text="Confirm", command=confirm_subfolders).pack(pady=10)

        # Wait until the window is closed
        subfolder_window.wait_window()

        if not all([test_case_name, selected_executions, burn_in_samples]):
            messagebox.showerror("Error", "All fields are required!")
            return

        # Ask the user to choose a directory where the structure will be created
        base_path = filedialog.askdirectory(title="Select Base Directory")
        if not base_path:
            messagebox.showerror("Error", "You must select a base directory!")
            return

        # Create the folder structure
        self.create_folder_structure(test_case_name, test_sample_names, selected_executions, burn_in_samples, selected_subfolders, base_path)

    def trigger_folder_creation(self):
        # Trigger the folder creation process
        self.get_user_input_and_create_folders()

    def load_directory(self):
        self.dir_path = filedialog.askdirectory()
        if self.dir_path:
            self.path_label.config(text="路徑："+self.dir_path)
            self.folder_tree.delete(*self.folder_tree.get_children())
            self.image_tree.delete(*self.image_tree.get_children())
            self.load_folders(self.dir_path)

    def load_folders(self, path, parent=""):
        for item in os.listdir(path):
            abs_path = os.path.join(path, item)
            if os.path.isdir(abs_path):
                node = self.folder_tree.insert(parent, 'end', text=item, open=False, tags=('folder',))
                self.load_folders(abs_path, node)
    def browse_check_file(self):
        self.dir_path = filedialog.askopenfilename()

    def on_folder_click(self, event):
        selected_item = self.folder_tree.selection()[0]
        folder_path = self.get_item_path(selected_item)
        self.load_files_and_images(folder_path)

    def load_files_and_images(self, path):
        self.image_tree.delete(*self.image_tree.get_children())
        for item in os.listdir(path):
            abs_path = os.path.join(path, item)
            if os.path.isfile(abs_path):
                file_extension = os.path.splitext(item)[1].lower()
                if file_extension in ['.png', '.jpg', '.jpeg']:
                    if abs_path not in self.image_cache:
                        img = Image.open(abs_path)
                        img.thumbnail((400, 400))
                        self.image_cache[abs_path] = ImageTk.PhotoImage(img)
                    image = self.image_cache[abs_path]
                    self.image_tree.insert('', 'end', text=item, image=image, tags=('image',))
                else:
                    self.image_tree.insert('', 'end', text=item, tags=('file',))

    def get_item_path(self, item_id):
        path_components = []
        while item_id:
            item_text = self.folder_tree.item(item_id, "text")
            path_components.append(item_text)
            item_id = self.folder_tree.parent(item_id)
        path_components.reverse()
        return os.path.join(self.dir_path, *path_components)

    def on_double_click(self, event):
        item_id = self.image_tree.selection()[0]
        item_text = self.image_tree.item(item_id, "text")
        item_path = self.get_image_item_path(item_id)
        original_extension = os.path.splitext(item_text)[1]

        new_name = CustomSimpledialog(self.root, title="Rename", message="請輸入新名稱：", initialvalue=item_text).result
        if new_name:
            new_extension = os.path.splitext(new_name)[1]
            if not new_extension:
                new_name += original_extension

            new_path = os.path.join(os.path.dirname(item_path), new_name)
            os.rename(item_path, new_path)
            self.image_tree.item(item_id, text=new_name)


    def get_image_item_path(self, item_id):
        # This method constructs the path based on the selected folder and image name.
        selected_folder = self.get_item_path(self.folder_tree.selection()[0])
        image_name = self.image_tree.item(item_id, "text")
        return os.path.join(selected_folder, image_name)

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
    def rename_images(self):
        folder_path = self.get_selected_folder_path()
        if not folder_path:
            messagebox.showwarning("警告", "請選擇一個資料夾！")
            return

        dir_name = os.path.basename(folder_path)
        patterns = self.default_naming_patterns.get(dir_name)

        if not patterns:
            patterns = self.ask_naming_patterns()

        if not patterns:
            messagebox.showwarning("警告", "無效的命名模式！")
            return

        burnin_patterns = self.burnin_patterns.get(dir_name, [])

        file_dict = self.create_file_dict(folder_path, patterns, burnin_patterns)
        pprint.pprint(file_dict)  # Print the dictionary to console

        if messagebox.askyesno("確認", "確定要重新命名這些圖像嗎？"):
            self.perform_rename(file_dict)
            messagebox.showinfo("完成", "所有圖像已重新命名。")

    def get_selected_folder_path(self):
        selected_item = self.folder_tree.selection()
        if not selected_item:
            return None
        return self.get_item_path(selected_item[0])

    def ask_naming_patterns(self):
        result = simpledialog.askstring("輸入模式", "請輸入命名模式 (例如: name1,name2;pattern1,pattern2):")
        if not result:
            return None

        try:
            parts = result.split(";")
            patterns = {parts[0]: parts[1].split(",")}
            return patterns
        except Exception as e:
            messagebox.showerror("錯誤", f"解析模式時出錯: {e}")
            return None

    def create_file_dict(self, folder_path, patterns, burnin_patterns):
        file_dict = {}
        for folder_name, sub_patterns in patterns.items():
            folder_full_path = os.path.join(folder_path, folder_name)
            if not os.path.exists(folder_full_path):
                os.makedirs(folder_full_path)

            for i, sub_pattern in enumerate(sub_patterns):
                file_dict[sub_pattern] = os.path.join(folder_full_path, f"{sub_pattern}.jpg")

            for burnin_pattern in burnin_patterns:
                burnin_folder = os.path.join(folder_full_path, "BurnIn")
                if not os.path.exists(burnin_folder):
                    os.makedirs(burnin_folder)
                file_dict[burnin_pattern] = os.path.join(burnin_folder, f"{burnin_pattern}.jpg")

        return file_dict

    def perform_rename(self, file_dict):
        for file_name, new_path in file_dict.items():
            old_path = os.path.join(self.get_selected_folder_path(), file_name + ".jpg")
            if os.path.exists(old_path):
                os.rename(old_path, new_path)
            else:
                print(f"File not found: {old_path}")

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
    def create_unique_name(self, directory, old_name, base_name, extension):
        if old_name == f"{base_name}{extension}":
            return old_name  # Return old name if it's already the same as the new name
        
        counter = 1
        new_name = f"{base_name}{extension}"
        new_file_path = os.path.join(directory, new_name)
        while os.path.exists(new_file_path) and old_name != new_name:
            new_name = f"{base_name}_{counter}{extension}"
            new_file_path = os.path.join(directory, new_name)
            counter += 1
        return new_name
    
    def check_and_rename_outer_layer(self):
        total_renamed = 0
        for test_type, subfolders in self.default_naming_patterns.items():
            for subfolder, naming_patterns in subfolders.items():
                folder_path = os.path.join(self.dir_path, test_type, subfolder)
                if not os.path.isdir(folder_path):
                    continue
                files = sorted(os.listdir(folder_path))
                images = [f for f in files if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
                if len(images) != len(naming_patterns):
                    messagebox.showwarning("Mismatch Warning", f"路徑：'{folder_path}'\n影像數目不對！")
                    continue
                for i, image_name in enumerate(images):
                    new_base_name = naming_patterns[i]
                    new_name = new_base_name + ".jpg"
                    file_name, file_extension = os.path.splitext(new_name)
                    src = os.path.join(folder_path, image_name)
                    dst = os.path.join(folder_path, new_name)
                    if os.path.exists(dst):
                        new_name = self.create_unique_name(folder_path, image_name, file_name, file_extension)
                        dst = os.path.join(folder_path, new_name)
                    os.rename(src, dst)
                    total_renamed += 1

        return total_renamed

    def prompt_burnin_subfolder_count(self):
        while True:
            user_input = simpledialog.askstring("Burn-in Subfolder Count", "請輸入燒機資料夾數量(輸入整數):")
            if user_input is None:
                return 0
            try:
                return int(user_input)
            except ValueError:
                messagebox.showerror("Error!!", "請輸入整數！")

    def check_and_rename_burnin_folders(self, burnin_subfolder_count):
        total_renamed = 0
        for test_type, subfolders in self.default_naming_patterns.items():
            for subfolder, _ in subfolders.items():
                burnin_folder_path = os.path.join(self.dir_path, test_type, subfolder, "BurnIn")
                if not os.path.isdir(burnin_folder_path):
                    continue
                burnin_subfolders = [f for f in os.listdir(burnin_folder_path) if os.path.isdir(os.path.join(burnin_folder_path, f)) and f.isdigit()]
                burnin_subfolders = sorted(burnin_subfolders, key=int)[:burnin_subfolder_count]

                if burnin_subfolder_count == 1:
                    images = [f for f in sorted(os.listdir(burnin_folder_path)) if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
                    naming_patterns = self.burnin_patterns.get(test_type, [])
                    if len(images) != len(naming_patterns):
                        messagebox.showwarning("Mismatch Warning", f"路徑：'{burnin_folder_path}' \n影像數目不對！請重新確認！")
                        continue
                    for i, image_name in enumerate(images):
                        new_base_name = naming_patterns[i]
                        new_name = new_base_name + ".jpg"
                        file_name, file_extension = os.path.splitext(new_name)
                        src = os.path.join(burnin_folder_path, image_name)
                        dst = os.path.join(burnin_folder_path, new_name)
                        if os.path.exists(dst):
                            new_name = self.create_unique_name(burnin_folder_path, image_name, file_name, file_extension)
                            dst = os.path.join(burnin_folder_path, new_name)
                        os.rename(src, dst)
                        total_renamed += 1
                else:
                    for subfolder_name in burnin_subfolders:
                        subfolder_path = os.path.join(burnin_folder_path, subfolder_name)
                        if not os.path.isdir(subfolder_path):
                            continue
                        images = [f for f in sorted(os.listdir(subfolder_path)) if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
                        naming_patterns = self.burnin_patterns.get(test_type, [])
                        if len(images) != len(naming_patterns):
                            messagebox.showwarning("Mismatch Warning", f"路徑：'{subfolder_path}'\n影像數目不對！請重新確認！")
                            continue
                        for j, image_name in enumerate(images):
                            new_base_name = naming_patterns[j]
                            new_name = new_base_name + ".jpg"
                            file_name, file_extension = os.path.splitext(new_name)
                            src = os.path.join(subfolder_path, image_name)
                            dst = os.path.join(subfolder_path, new_name)
                            if os.path.exists(dst):
                                new_name = self.create_unique_name(subfolder_path, image_name, file_name, file_extension)
                                dst = os.path.join(subfolder_path, new_name)
                            os.rename(src, dst)
                            total_renamed += 1

        return total_renamed

    def rename_images(self):
        total_renamed_outer = self.check_and_rename_outer_layer()
        burnin_subfolder_count = self.prompt_burnin_subfolder_count()
        total_renamed_burnin = self.check_and_rename_burnin_folders(burnin_subfolder_count)
        messagebox.showinfo("Rename Complete", f"Renamed {total_renamed_outer + total_renamed_burnin} images.")
class CustomSimpledialog(simpledialog.Dialog):
    def __init__(self, parent, title, message, initialvalue=""):
        self.message = message
        self.initialvalue = initialvalue
        super().__init__(parent, title)

    def body(self, master):
        tk.Label(master, text=self.message).pack()
        self.entry = tk.Entry(master)
        self.entry.pack(padx=20)
        self.entry.insert(0, self.initialvalue)
        return self.entry

    def apply(self):
        self.result = self.entry.get()

if __name__ == "__main__":
    root = tk.Tk()
    app = FileExplorerApp(root)
    root.mainloop()