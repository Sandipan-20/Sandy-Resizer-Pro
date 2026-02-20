import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk, ImageDraw
import os
import threading
import glob
import io
import math
import queue

# ===========================import Drag and Drop support=========================
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES

    DND_SUPPORT = True
except ImportError:
    DND_SUPPORT = False

# Constants
SUPPORTED_FORMATS = ['.bmp', '.jpg', '.jpeg', '.gif', '.png', '.pdf', '.webp', '.ico']


# ========================================== LOGO GENERATOR ======================================
def create_sunflower_image(size=(64, 64)):
    img = Image.new("RGBA", size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    w, h = size
    cx, cy = w // 2, h // 2
    petal_count = 16
    petal_length = w // 2.5
    petal_width = w // 8

    for i in range(petal_count):
        angle = (360 / petal_count) * i
        petal = Image.new("RGBA", size, (0, 0, 0, 0))
        p_draw = ImageDraw.Draw(petal)
        p_draw.ellipse([(cx - petal_width, cy - petal_length - 5), (cx + petal_width, cy + 5)], fill=(255, 215, 0))
        rotated = petal.rotate(angle, center=(cx, cy), resample=Image.BICUBIC)
        img.paste(rotated, (0, 0), rotated)

    center_radius = w // 4
    draw.ellipse([(cx - center_radius, cy - center_radius), (cx + center_radius, cy + center_radius)],
                 fill=(101, 67, 33))
    for i in range(0, 360, 30):
        r = center_radius // 2
        dx = int(r * math.cos(math.radians(i)))
        dy = int(r * math.sin(math.radians(i)))
        draw.ellipse([(cx + dx - 1, cy + dy - 1), (cx + dx + 1, cy + dy + 1)], fill=(60, 40, 20))
    return img


# !--- PRESETS ---
PRESETS_GENERAL = {
    "Custom": (0, 0), "1:1 (Square - 1080x1080)": (1080, 1080),
    "4:3 (Standard - 1440x1080)": (1440, 1080), "3:2 (Photo - 1440x960)": (1440, 960),
    "16:9 (Widescreen - 1920x1080)": (1920, 1080), "21:9 (Ultra-Wide - 2560x1080)": (2560, 1080),
    "Passport Photo (35x45 mm)": (35, 45), "HD Wallpaper (1920x1080)": (1920, 1080),
    "4K Wallpaper (3840x2160)": (3840, 2160), "Instagram Post (1080x1080)": (1080, 1080),
}
PRESETS_PDF = {
    "Custom": (0, 0), "A4 (210x297 mm)": (210, 297), "A3 (297x420 mm)": (297, 420),
    "A5 (148x210 mm)": (148, 210), "Letter (216x279 mm)": (216, 279),
}
PRESETS_ICO = {
    "Custom": (0, 0), "16x16 px": (16, 16), "32x32 px": (32, 32),
    "64x64 px": (64, 64), "256x256 px": (256, 256),
}


class SandyResizerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sandy Resizer Pro")
        self.root.geometry("1100x680")
        self.root.minsize(800, 600)

        # Set Logo
        self.logo_image = create_sunflower_image((64, 64))
        self.logo_photo = ImageTk.PhotoImage(self.logo_image)
        self.root.iconphoto(True, self.logo_photo)

        self.image_list = []
        self.thumbnails = []

        self.current_preview_index = 0
        self.preview_running = False

        self.current_orig_w = 0
        self.current_orig_h = 0
        self.orig_img_ratio = 0

        # Cache for responsive resizing
        self.cached_preview_img = None

        self.loading_queue = queue.Queue()
        self.is_loading = False

        self.setup_ui()
        self.load_ui_settings()

        self.root.after(100, self.start_preview_loop)

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- Left Panel ---
        left_frame = ttk.Frame(main_frame, width=250)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))

        # Add Buttons
        btn_frame = ttk.Frame(left_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 5))

        self.btn_add_files = ttk.Button(btn_frame, text="Add Files", command=self.add_files)
        self.btn_add_files.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)

        self.btn_add_folder = ttk.Button(btn_frame, text="Add Folder", command=self.add_folder)
        self.btn_add_folder.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)

        # Status Label
        self.status_var = tk.StringVar(value="")
        self.status_label = ttk.Label(left_frame, textvariable=self.status_var, foreground="blue",
                                      font=("Arial", 9, "italic"))
        self.status_label.pack(fill=tk.X, pady=2)

        # Treeview for Files
        tree_frame = ttk.Frame(left_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        self.file_tree = ttk.Treeview(tree_frame, selectmode='extended', show='tree')
        self.file_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.file_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_tree.configure(yscrollcommand=scrollbar.set)

        self.file_tree.column("#0", width=230, minwidth=200)
        self.file_tree.bind('<<TreeviewSelect>>', self.on_tree_select)

        # Remove Buttons Frame
        remove_btn_frame = ttk.Frame(left_frame)
        remove_btn_frame.pack(fill=tk.X, pady=5)

        ttk.Button(remove_btn_frame, text="Remove", command=self.remove_selected).pack(side=tk.LEFT, expand=True,
                                                                                       fill=tk.X, padx=(0, 2))
        ttk.Button(remove_btn_frame, text="Remove All", command=self.remove_all).pack(side=tk.LEFT, expand=True,
                                                                                      fill=tk.X, padx=(2, 0))

        if DND_SUPPORT:
            self.file_tree.drop_target_register(DND_FILES)
            self.file_tree.dnd_bind('<<Drop>>', self.on_drop)

        # --- Center Panel (Responsive) ---
        center_frame = ttk.Frame(main_frame)
        center_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.preview_canvas = tk.Canvas(center_frame, bg="#2b2b2b", highlightthickness=1, highlightbackground="black")
        self.preview_canvas.pack(fill=tk.BOTH, expand=True)

        # Bind resize event to canvas
        self.preview_canvas.bind("<Configure>", self.on_canvas_resize)

        self.preview_info_var = tk.StringVar(value="Add images to begin")
        ttk.Label(center_frame, textvariable=self.preview_info_var).pack()

        # --- Right Panel ---
        right_frame = ttk.Frame(main_frame, width=320)
        right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(10, 0))
        right_frame.pack_propagate(False)  # Keep width fixed

        # 1. OUTPUT SETTINGS
        self.file_frame = ttk.LabelFrame(right_frame, text="Output Settings", padding="10")
        self.file_frame.pack(fill=tk.X, pady=5)

        ttk.Label(self.file_frame, text="Format:").grid(row=0, column=0, sticky='w')
        self.format_var = tk.StringVar(value="JPEG")
        self.format_combo = ttk.Combobox(self.file_frame, textvariable=self.format_var,
                                         values=["JPEG", "PNG", "WEBP", "GIF", "BMP", "ICO", "PDF"], state="readonly")
        self.format_combo.grid(row=0, column=1, sticky='ew')
        self.format_combo.bind("<<ComboboxSelected>>", self.on_format_change)

        ttk.Label(self.file_frame, text="Rotate:").grid(row=1, column=0, sticky='w', pady=2)
        self.rotate_var = tk.StringVar(value="0")
        ttk.Combobox(self.file_frame, textvariable=self.rotate_var, values=["0", "90", "180", "270"],
                     state="readonly").grid(row=1, column=1, sticky='ew', pady=2)

        # 2. RESIZE DIMENSIONS
        self.dim_frame = ttk.LabelFrame(right_frame, text="Resize Dimensions", padding="10")
        self.dim_frame.pack(fill=tk.X, pady=5)

        # Grid Configuration for Alignment
        self.dim_frame.columnconfigure(1, weight=1)

        # Row 0: Presets
        ttk.Label(self.dim_frame, text="Presets:").grid(row=0, column=0, sticky='w')
        self.preset_var = tk.StringVar(value="Custom")
        self.preset_combo = ttk.Combobox(self.dim_frame, textvariable=self.preset_var,
                                         values=list(PRESETS_GENERAL.keys()), state="readonly")
        self.preset_combo.grid(row=0, column=1, columnspan=2, sticky='ew')
        self.preset_combo.bind("<<ComboboxSelected>>", self.on_preset_change)

        # Row 1: Unit
        ttk.Label(self.dim_frame, text="Unit:").grid(row=1, column=0, sticky='w', pady=5)
        self.unit_var = tk.StringVar(value="percent")
        self.unit_menu = ttk.Combobox(self.dim_frame, textvariable=self.unit_var,
                                      values=["px", "percent", "inch", "cm", "mm"], state="readonly")
        self.unit_menu.grid(row=1, column=1, columnspan=2, sticky='ew', pady=5)
        self.unit_menu.bind("<<ComboboxSelected>>", self.on_unit_change)

        # Row 2: DPI (Hidden for px/percent)
        self.dpi_label = ttk.Label(self.dim_frame, text="DPI:")
        self.dpi_label.grid(row=2, column=0, sticky='w', pady=2)

        self.dpi_var = tk.StringVar(value="96")
        self.dpi_combo = ttk.Combobox(self.dim_frame, textvariable=self.dpi_var,
                                      values=["72", "96", "150", "300", "600"], state="readonly")
        self.dpi_combo.grid(row=2, column=1, sticky='ew', pady=2)
        self.dpi_combo.bind("<<ComboboxSelected>>", self.on_dpi_change)

        # Row 3: Width/Height Frame
        self.wh_frame = ttk.Frame(self.dim_frame)
        # Internal WH Frame Grid
        ttk.Label(self.wh_frame, text="Width:").grid(row=0, column=0, sticky='w')
        self.width_var = tk.StringVar()
        self.width_entry = ttk.Entry(self.wh_frame, textvariable=self.width_var, width=10)
        self.width_entry.grid(row=0, column=1, sticky='ew')
        self.width_entry.bind("<KeyRelease>", self.on_width_change)

        ttk.Label(self.wh_frame, text="Height:").grid(row=1, column=0, sticky='w')
        self.height_var = tk.StringVar()
        self.height_entry = ttk.Entry(self.wh_frame, textvariable=self.height_var, width=10)
        self.height_entry.grid(row=1, column=1, sticky='ew')
        self.height_entry.bind("<KeyRelease>", self.on_height_change)

        # Row 3: Percent Frame
        self.percent_frame = ttk.Frame(self.dim_frame)
        self.percent_var = tk.IntVar(value=100)
        self.percent_label_var = tk.StringVar(value="100%")
        ttk.Label(self.percent_frame, text="Scale:").pack(side=tk.LEFT, padx=5)
        self.percent_slider = ttk.Scale(self.percent_frame, from_=1, to=100, variable=self.percent_var,
                                        orient=tk.HORIZONTAL, command=self.on_percent_change)
        self.percent_slider.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.percent_slider.bind("<ButtonRelease-1>", self.on_slider_release)
        ttk.Label(self.percent_frame, textvariable=self.percent_label_var, width=5).pack(side=tk.LEFT)

        # Row 4: Checkbox
        self.keep_ratio_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(self.dim_frame, text="Maintain Original Ratio", variable=self.keep_ratio_var).grid(row=4,
                                                                                                           column=0,
                                                                                                           columnspan=3,
                                                                                                           sticky='w',
                                                                                                           pady=5)

        # 3.  !-----------------------------------OUTPUT QUALITY-----------------------------------
        self.quality_frame = ttk.LabelFrame(right_frame, text="Output Quality", padding="10")
        self.quality_var = tk.IntVar(value=80)
        self.quality_label_var = tk.StringVar(value="80%")
        ttk.Scale(self.quality_frame, from_=1, to=100, variable=self.quality_var, orient=tk.HORIZONTAL,
                  command=self.on_quality_change).pack(fill=tk.X)
        q_label_frame = ttk.Frame(self.quality_frame)
        q_label_frame.pack(fill=tk.X)
        ttk.Label(q_label_frame, text="Low (Small)").pack(side=tk.LEFT)
        ttk.Label(q_label_frame, textvariable=self.quality_label_var).pack(side=tk.LEFT, expand=True)
        ttk.Label(q_label_frame, text="High (Large)").pack(side=tk.RIGHT)

        # 4. -----------------------------------SIZE ESTIMATOR-----------------------------------
        calc_frame = ttk.LabelFrame(right_frame, text="Size Estimator", padding="10")
        calc_frame.pack(fill=tk.X, pady=5)
        self.calculated_size_var = tk.StringVar(value="Size: --")
        ttk.Label(calc_frame, textvariable=self.calculated_size_var, font=("Arial", 10, "bold")).pack(pady=5)
        ttk.Button(calc_frame, text="Check Size (Process)", command=self.calculate_buffer_size).pack(fill=tk.X)

        # 5. -----------------------------------FILENAME SETTINGS-----------------------------------
        self.name_frame = ttk.LabelFrame(right_frame, text="Filename Settings", padding="10")
        self.name_frame.pack(fill=tk.X, pady=5)

        self.name_frame.columnconfigure(1, weight=1)

        ttk.Label(self.name_frame, text="Rename:").grid(row=0, column=0, sticky='w', padx=5)
        self.rename_var = tk.StringVar(value="[Original Name]_[Width]×[Height]")
        self.rename_combo = ttk.Combobox(self.name_frame, textvariable=self.rename_var,
                                         values=["Original Name", "[Original Name]_[Width]×[Height]", "Add Suffix"],
                                         state="readonly")
        self.rename_combo.grid(row=0, column=1, sticky='ew', padx=5)
        self.rename_combo.bind("<<ComboboxSelected>>", self.on_rename_option_change)

        # !-----------------------------------Suffix Input Widgets-----------------------------------
        self.suffix_var = tk.StringVar()
        self.suffix_label = ttk.Label(self.name_frame, text="Suffix:")
        self.suffix_entry = ttk.Entry(self.name_frame, textvariable=self.suffix_var)

        # 6. ACTIONS
        action_frame = ttk.Frame(right_frame)
        action_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=10)

        ttk.Button(action_frame, text="Resize", command=self.resize_single).pack(fill=tk.X, pady=2)
        ttk.Button(action_frame, text="Resize All", command=self.start_resize_thread).pack(fill=tk.X, pady=2)

        self.toggle_quality_visibility()
        self.toggle_percent_ui()
        self.on_rename_option_change(None)

    # --------------------------------------Responsive Canvas Logic --------------------------------------

    def on_canvas_resize(self, event):
        """Called when canvas is resized. Redraws image to fit new size."""
        if self.cached_preview_img:
            self.draw_preview_image()

    def draw_preview_image(self):
        """Draws the cached image onto the canvas, fitting it."""
        if not self.cached_preview_img:
            return

        canvas_w = self.preview_canvas.winfo_width()
        canvas_h = self.preview_canvas.winfo_height()

        if canvas_w < 10 or canvas_h < 10: return

        preview_img = self.cached_preview_img.copy()
        preview_img.thumbnail((canvas_w, canvas_h), Image.Resampling.LANCZOS)

        self.tk_img = ImageTk.PhotoImage(preview_img)
        self.preview_canvas.delete("all")
        self.preview_canvas.create_image(canvas_w // 2, canvas_h // 2, image=self.tk_img)

    # !-------------------------------------- Filename Logic --------------------------------------

    def on_rename_option_change(self, event):
        if self.rename_var.get() == "Add Suffix":
            self.suffix_label.grid(row=1, column=0, sticky='w', padx=5, pady=(5, 0))
            self.suffix_entry.grid(row=1, column=1, sticky='ew', padx=5, pady=(5, 0))
        else:
            self.suffix_label.grid_forget()
            self.suffix_entry.grid_forget()

    def get_output_filename(self, base_name, w, h):
        choice = self.rename_var.get()
        if choice == "[Original Name]_[Width]×[Height]":
            return f"{base_name}_{w}x{h}"
        elif choice == "Add Suffix":
            suffix = self.suffix_var.get()
            return f"{base_name}_{suffix}"
        else:
            return base_name

    # !-------------------------------------- Core Resize Logic Helpers --------------------------------------

    def get_current_settings(self):
        try:
            dpi_val = float(self.dpi_var.get())
        except ValueError:
            dpi_val = 96.0

        return {
            'width': self.width_var.get(),
            'height': self.height_var.get(),
            'unit': self.unit_var.get(),
            'keep_ratio': self.keep_ratio_var.get(),
            'format': self.format_var.get(),
            'rotate': self.rotate_var.get(),
            'quality': self.quality_var.get(),
            'rename_option': self.rename_var.get(),
            'suffix': self.suffix_var.get(),
            'dpi': dpi_val
        }

    def calculate_target_dimensions(self, img, settings):
        try:
            orig_w, orig_h = img.size
            unit = settings['unit']

            try:
                target_w = float(settings['width'])
            except:
                target_w = 0

            try:
                target_h = float(settings['height'])
            except:
                target_h = 0

            if unit == "percent":
                scale = (target_w if target_w > 0 else 100) / 100.0
                new_w = int(orig_w * scale)
                new_h = int(orig_h * scale)

            elif unit in ["inch", "cm", "mm"]:
                val_w, val_h = target_w, target_h
                dpi = settings['dpi']

                # -----------------------------------Safety check for DPI-----------------------------------
                if dpi <= 0: dpi = 96.0

                if unit == "cm":
                    val_w /= 2.54;
                    val_h /= 2.54
                elif unit == "mm":
                    val_w /= 25.4;
                    val_h /= 25.4

                new_w = int(val_w * dpi) if target_w > 0 else 0
                new_h = int(val_h * dpi) if target_h > 0 else 0

                if settings['keep_ratio']:
                    if target_w > 0 and target_h <= 0:
                        new_h = int(new_w * (orig_h / orig_w))
                    elif target_h > 0 and target_w <= 0:
                        new_w = int(new_h * (orig_w / orig_h))

            else:
                # -----------------------------------Pixels-----------------------------------
                new_w = int(target_w) if target_w > 0 else orig_w
                new_h = int(target_h) if target_h > 0 else orig_h

                if settings['keep_ratio']:
                    if target_w > 0 and target_h <= 0:
                        new_h = int(new_w * (orig_h / orig_w))
                    elif target_h > 0 and target_w <= 0:
                        new_w = int(new_h * (orig_w / orig_h))

            # !-----------------------------------Safety check for extremely large dimensions
            MAX_DIM = 20000
            if new_w > MAX_DIM: new_w = MAX_DIM
            if new_h > MAX_DIM: new_h = MAX_DIM

            if new_w <= 0: new_w = orig_w
            if new_h <= 0: new_h = orig_h

            return max(1, new_w), max(1, new_h)
        except Exception as e:
            print(f"Dimension Error: {e}")
            return img.size

    # -------------------------------------- Action Methods --------------------------------------

    def resize_single(self):
        if not self.image_list:
            messagebox.showwarning("No Selection", "Please select a photo to resize.")
            return

        path = self.image_list[self.current_preview_index]

        try:
            img = Image.open(path)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open image: {e}")
            return

        settings = self.get_current_settings()

        angle = int(settings['rotate'])
        if angle != 0: img = img.rotate(-angle, expand=True)

        w, h = self.calculate_target_dimensions(img, settings)

        base_name = os.path.splitext(os.path.basename(path))[0]
        final_name = self.get_output_filename(base_name, w, h)

        fmt = settings['format']
        ext_map = {
            "JPEG": [".jpg", ".jpeg"], "PNG": [".png"], "WEBP": [".webp"],
            "GIF": [".gif"], "BMP": [".bmp"], "ICO": [".ico"], "PDF": [".pdf"]
        }
        ext_list = ext_map.get(fmt, [".jpg"])

        save_path = filedialog.asksaveasfilename(
            initialfile=final_name,
            defaultextension=ext_list[0],
            filetypes=[(f"{fmt} Files", f"*{ext_list[0]}")]
        )

        if not save_path: return

        threading.Thread(target=self._process_single, args=(path, save_path, settings, w, h), daemon=True).start()

    def _process_single(self, input_path, output_path, settings, w, h):
        try:
            img = Image.open(input_path)

            angle = int(settings['rotate'])
            if angle != 0: img = img.rotate(-angle, expand=True)

            img = img.resize((w, h), Image.Resampling.LANCZOS)

            save_args = {}
            fmt = settings['format']

            # Determine DPI based on unit-----------------------------------
            # -----------------------------------Force 96 DPI for PX and Percent, otherwise use settings
            if settings['unit'] in ['px', 'percent']:
                dpi = 96.0
            else:
                dpi = settings['dpi']

            if fmt == "JPEG" and img.mode in ("RGBA", "P"): img = img.convert("RGB")

            # -----------------------------------DPI Handling for standard formats-----------------------------------
            if fmt in ["JPEG", "PNG", "BMP", "WEBP"]:
                save_args['dpi'] = (dpi, dpi)

            if fmt in ["JPEG", "WEBP"]:
                save_args['quality'] = settings['quality']
            if fmt == "PNG":
                save_args['compress_level'] = int(9 - (settings['quality'] / 100) * 9)

            if fmt == "PDF":
                if img.mode in ("RGBA", "P"): img = img.convert("RGB")
                # ------------------------------Use user-defined DPI for PDF resolution-----------------------------
                img.save(output_path, "PDF", resolution=dpi)
            else:
                img.save(output_path, **save_args)

            self.root.after(0, lambda: messagebox.showinfo("Done", f"Image saved to:\n{output_path}"))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", f"Could not save image:\n{e}"))

    def start_resize_thread(self):
        if not self.image_list:
            messagebox.showwarning("No Files", "Please add files to resize.")
            return

        output_dir = filedialog.askdirectory(title="Select Output Folder")
        if not output_dir: return

        settings = self.get_current_settings()
        threading.Thread(target=self.resize_all, args=(output_dir, settings), daemon=True).start()

    def resize_all(self, output_dir, settings):
        success_count, errors = 0, 0
        fmt_choice = settings['format']

        # -----------------------------------Determine DPI based on unit-----------------------------------
        if settings['unit'] in ['px', 'percent']:
            dpi = 96.0
        else:
            dpi = settings['dpi']

        for path in self.image_list:
            try:
                if path.lower().endswith('.pdf'):
                    errors += 1
                    continue

                img = Image.open(path)

                angle = int(settings['rotate'])
                if angle != 0: img = img.rotate(-angle, expand=True)

                w, h = self.calculate_target_dimensions(img, settings)
                img = img.resize((w, h), Image.Resampling.LANCZOS)

                base_name = os.path.splitext(os.path.basename(path))[0]
                final_name = self.get_output_filename(base_name, w, h)

                ext = f".{fmt_choice.lower()}"
                if ext == ".jpeg": ext = ".jpg"

                out_path = os.path.join(output_dir, f"{final_name}{ext}")

                counter = 1
                while os.path.exists(out_path):
                    out_path = os.path.join(output_dir, f"{final_name}_{counter}{ext}")
                    counter += 1

                save_args = {}
                if fmt_choice in ["JPEG", "JPG"]:
                    if img.mode in ("RGBA", "P"): img = img.convert("RGB")
                    save_args['quality'] = settings['quality']

                # -----------------------------------Apply DPI to standard formats-----------------------------------
                if fmt_choice in ["JPEG", "JPG", "PNG", "BMP", "WEBP"]:
                    save_args['dpi'] = (dpi, dpi)

                if fmt_choice == "WEBP":
                    save_args['quality'] = settings['quality']
                if fmt_choice == "PNG":
                    save_args['compress_level'] = int(9 - (settings['quality'] / 100) * 9)

                if fmt_choice == "PDF":
                    if img.mode in ("RGBA", "P"): img = img.convert("RGB")
                    img.save(out_path, "PDF", resolution=dpi)
                else:
                    img.save(out_path, **save_args)

                success_count += 1
            except Exception as e:
                print(f"Error processing {path}: {e}")
                errors += 1

        msg = f"Completed!\nSuccess: {success_count}\nErrors/Skipped: {errors}"
        self.root.after(0, lambda: messagebox.showinfo("Done", msg))

    # -------------------------------------- Loading Logic --------------------------------------

    def set_ui_loading(self, loading_state):
        self.is_loading = loading_state
        if loading_state:
            self.btn_add_files.config(state='disabled')
            self.btn_add_folder.config(state='disabled')
            self.status_var.set("Loading...")
            self.root.after(50, self.process_loading_queue)
        else:
            self.btn_add_files.config(state='normal')
            self.btn_add_folder.config(state='normal')
            self.status_var.set("")

    def threaded_load_files(self, files):
        for f in files:
            ext = os.path.splitext(f)[1].lower()
            if ext in SUPPORTED_FORMATS:
                pil_thumb = None
                if not f.lower().endswith('.pdf'):
                    try:
                        img = Image.open(f)
                        img.thumbnail((32, 32), Image.Resampling.LANCZOS)
                        pil_thumb = img
                    except:
                        pass
                self.loading_queue.put((f, os.path.basename(f), pil_thumb))
        self.loading_queue.put(None)

    def process_loading_queue(self):
        if not self.is_loading: return
        try:
            while True:
                item = self.loading_queue.get_nowait()
                if item is None:
                    self.set_ui_loading(False)
                    return
                else:
                    f, name, pil_thumb = item
                    if f not in self.image_list:
                        self.image_list.append(f)
                        thumb_photo = None
                        if pil_thumb:
                            thumb_photo = ImageTk.PhotoImage(pil_thumb)
                            self.thumbnails.append(thumb_photo)
                        else:
                            self.thumbnails.append(None)

                        if thumb_photo:
                            self.file_tree.insert('', 'end', text=name, image=thumb_photo)
                        else:
                            self.file_tree.insert('', 'end', text=name)
                        self.status_var.set(f"Loaded {len(self.image_list)} files...")
        except queue.Empty:
            if self.is_loading:
                self.root.after(100, self.process_loading_queue)

    # ----------------------------------- UI Logic Methods -----------------------------------------

    def add_files(self):
        files = filedialog.askopenfilenames(
            filetypes=[("Image Files", "*.bmp *.jpg *.jpeg *.gif *.png *.pdf *.webp *.ico")])
        if files:
            self.set_ui_loading(True)
            threading.Thread(target=self.threaded_load_files, args=(files,), daemon=True).start()

    def add_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            files = []
            for ext in SUPPORTED_FORMATS:
                files.extend(glob.glob(os.path.join(folder, f"*{ext}")))
                files.extend(glob.glob(os.path.join(folder, f"*{ext.upper()}")))
            if files:
                self.set_ui_loading(True)
                threading.Thread(target=self.threaded_load_files, args=(files,), daemon=True).start()

    def on_drop(self, event):
        files = self.root.tk.splitlist(event.data)
        if files:
            self.set_ui_loading(True)
            threading.Thread(target=self.threaded_load_files, args=(files,), daemon=True).start()

    def on_percent_change(self, value):
        val = int(float(value))
        self.percent_label_var.set(f"{val}%")
        self.width_var.set(str(val))

    def on_slider_release(self, event):
        self.update_preview()

    def on_dpi_change(self, event):
        self.update_preview()

    def on_unit_change(self, event):
        # --------------------------------Convert values if image loaded-----------------------------------
        if self.current_orig_w > 0:
            unit = self.unit_var.get()
            try:
                dpi = float(self.dpi_var.get())
            except:
                dpi = 96.0

            w_inches = self.current_orig_w / dpi
            h_inches = self.current_orig_h / dpi

            if unit == "px":
                self.width_var.set(str(self.current_orig_w))
                self.height_var.set(str(self.current_orig_h))
            elif unit == "inch":
                self.width_var.set(str(round(w_inches, 2)))
                self.height_var.set(str(round(h_inches, 2)))
            elif unit == "cm":
                self.width_var.set(str(round(w_inches * 2.54, 2)))
                self.height_var.set(str(round(h_inches * 2.54, 2)))
            elif unit == "mm":
                self.width_var.set(str(round(w_inches * 25.4, 2)))
                self.height_var.set(str(round(h_inches * 25.4, 2)))

        self.toggle_percent_ui()
        self.update_preview()

    def toggle_percent_ui(self):
        unit = self.unit_var.get()

        # Handle DPI visibility
        # Show DPI only for physical units
        if unit in ["inch", "cm", "mm"]:
            self.dpi_label.grid(row=2, column=0, sticky='w', pady=2)
            self.dpi_combo.grid(row=2, column=1, sticky='ew', pady=2)
        else:
            # Hide for px and percent
            self.dpi_label.grid_remove()
            self.dpi_combo.grid_remove()

        # Handle Slider vs Entry (Row 3)
        if unit == "percent":
            self.wh_frame.grid_remove()
            self.percent_frame.grid(row=3, column=0, columnspan=3, sticky='ew')
            self.keep_ratio_var.set(True)
        else:
            self.percent_frame.grid_remove()
            self.wh_frame.grid(row=3, column=0, columnspan=3, sticky='ew')

    def on_format_change(self, event):
        self.update_preset_list()
        self.toggle_quality_visibility()
        self.update_preview()

    def update_preset_list(self):
        fmt = self.format_var.get()
        if fmt == "PDF":
            presets = PRESETS_PDF
            self.unit_var.set("mm")
        elif fmt == "ICO":
            presets = PRESETS_ICO
            self.unit_var.set("px")
        else:
            presets = PRESETS_GENERAL
            self.unit_var.set("percent")

        self.preset_combo['values'] = list(presets.keys())
        self.preset_var.set("Custom")
        self.toggle_percent_ui()

    def toggle_quality_visibility(self):
        selected_format = self.format_var.get()
        self.quality_frame.pack_forget()
        if selected_format in ["JPEG", "WEBP"]:
            self.quality_frame.pack(fill=tk.X, pady=5, after=self.dim_frame)

    def on_quality_change(self, value):
        val = int(float(value))
        self.quality_label_var.set(f"{val}%")

    def remove_selected(self):
        selected_items = self.file_tree.selection()
        children = self.file_tree.get_children('')
        indices_to_remove = []

        for item in selected_items:
            try:
                idx = children.index(item)
                indices_to_remove.append(idx)
            except ValueError:
                pass

        for idx in sorted(indices_to_remove, reverse=True):
            self.image_list.pop(idx)
            self.thumbnails.pop(idx)

        for item in selected_items:
            self.file_tree.delete(item)

        if not self.image_list:
            self.cached_preview_img = None
            self.preview_canvas.delete("all")
            self.preview_info_var.set("No images loaded")
            self.calculated_size_var.set("Size: --")
        else:
            self.cached_preview_img = None
            self.preview_canvas.delete("all")
            self.preview_info_var.set("Select an image")

    def remove_all(self):
        self.image_list = []
        self.thumbnails = []
        self.cached_preview_img = None
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)

        self.preview_canvas.delete("all")
        self.preview_info_var.set("No images loaded")
        self.calculated_size_var.set("Size: --")
        self.current_preview_index = 0

    def on_tree_select(self, event):
        selection = self.file_tree.selection()
        if selection:
            children = self.file_tree.get_children('')
            try:
                idx = children.index(selection[0])
                self.current_preview_index = idx
                self.preview_running = False
                self.update_preview(update_ratio=True)
            except ValueError:
                pass

    def on_preset_change(self, event):
        selected = self.preset_var.get()
        fmt = self.format_var.get()
        if fmt == "PDF":
            presets = PRESETS_PDF
        elif fmt == "ICO":
            presets = PRESETS_ICO
        else:
            presets = PRESETS_GENERAL

        if selected != "Custom":
            w, h = presets[selected]
            if "mm" in selected:
                self.unit_var.set("mm")
            else:
                self.unit_var.set("px")

            self.width_var.set(str(w))
            self.height_var.set(str(h))

            self.keep_ratio_var.set(False)
            self.toggle_percent_ui()
            self.update_preview()

    def on_width_change(self, event):
        self.preset_var.set("Custom")
        if not self.keep_ratio_var.get() or self.orig_img_ratio == 0: return
        try:
            w = float(self.width_var.get())
            new_h = w * self.orig_img_ratio
            if self.unit_var.get() in ["px"]:
                self.height_var.set(str(int(new_h)))
            else:
                self.height_var.set(str(round(new_h, 2)))
        except ValueError:
            pass
        self.update_preview()

    def on_height_change(self, event):
        self.preset_var.set("Custom")
        if not self.keep_ratio_var.get() or self.orig_img_ratio == 0: return
        try:
            h = float(self.height_var.get())
            new_w = h / self.orig_img_ratio
            if self.unit_var.get() in ["px"]:
                self.width_var.set(str(int(new_w)))
            else:
                self.width_var.set(str(round(new_w, 2)))
        except ValueError:
            pass
        self.update_preview()

    def update_preview(self, update_ratio=False):
        if not self.image_list: return
        try:
            path = self.image_list[self.current_preview_index]
            if path.lower().endswith('.pdf'):
                self.preview_canvas.delete("all")
                self.preview_canvas.create_text(self.preview_canvas.winfo_width() // 2,
                                                self.preview_canvas.winfo_height() // 2,
                                                text="PDF File\n(Preview not available)", fill="white",
                                                justify='center')
                self.preview_info_var.set(f"File: {os.path.basename(path)} (PDF)")
                self.cached_preview_img = None
                return

            img = Image.open(path)

            if update_ratio:
                self.current_orig_w, self.current_orig_h = img.size
                if self.current_orig_w > 0: self.orig_img_ratio = self.current_orig_h / self.current_orig_w

            angle = int(self.rotate_var.get())
            if angle != 0:
                img = img.rotate(-angle, expand=True)
                if update_ratio: self.orig_img_ratio = img.size[1] / img.size[0]

            self.cached_preview_img = img

            settings = self.get_current_settings()
            w, h = self.calculate_target_dimensions(img, settings)

            self.draw_preview_image()

            self.preview_info_var.set(f"File: {os.path.basename(path)}\nNew Size: {w} x {h} px")
        except Exception:
            pass

    def start_preview_loop(self):
        if self.preview_running and self.image_list:
            children = self.file_tree.get_children()
            if children:
                next_idx = (self.current_preview_index + 1) % len(self.image_list)
                self.current_preview_index = next_idx
                self.file_tree.selection_set(children[next_idx])
                self.file_tree.see(children[next_idx])
                self.update_preview(update_ratio=True)
        self.root.after(2000, self.start_preview_loop)

    def calculate_buffer_size(self):
        if not self.image_list:
            messagebox.showwarning("No Files", "Please add files first.")
            return
        try:
            path = self.image_list[self.current_preview_index]
            if path.lower().endswith('.pdf'):
                self.calculated_size_var.set("Size: PDF (N/A)")
                return
            img = Image.open(path)
            angle = int(self.rotate_var.get())
            if angle != 0: img = img.rotate(-angle, expand=True)

            settings = self.get_current_settings()
            w, h = self.calculate_target_dimensions(img, settings)

            img = img.resize((w, h), Image.Resampling.LANCZOS)
            save_fmt = self.format_var.get()
            if save_fmt == "JPEG" and img.mode in ("RGBA", "P"): img = img.convert("RGB")
            buffer = io.BytesIO()
            quality = self.quality_var.get()
            params = {}
            if save_fmt in ["JPEG", "WEBP"]:
                params['quality'] = quality
            elif save_fmt == "PNG":
                params['compress_level'] = int(9 - (quality / 100) * 9)
            img.save(buffer, format=save_fmt, **params)
            size_bytes = buffer.tell()
            if size_bytes < 1024:
                size_str = f"{size_bytes} Bytes"
            elif size_bytes < 1024 * 1024:
                size_str = f"{size_bytes / 1024:.2f} KB"
            else:
                size_str = f"{size_bytes / (1024 * 1024):.2f} MB"
            self.calculated_size_var.set(f"Size: {size_str} ({save_fmt})")
        except Exception as e:
            print(f"Error calculating size: {e}")
            self.calculated_size_var.set("Error calculating size")

    def load_ui_settings(self):
        self.update_preset_list()
        self.toggle_quality_visibility()
        self.toggle_percent_ui()
        self.on_rename_option_change(None)


if __name__ == "__main__":
    if DND_SUPPORT:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    app = SandyResizerApp(root)
    root.mainloop()