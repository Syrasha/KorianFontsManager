import os
import json
import ctypes
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font as tkfont, simpledialog, colorchooser
from PIL import ImageFont, Image, ImageTk
import threading

# Font loading logic for Windows
def load_font(font_path):
    try:
        # FR_PRIVATE = 0x10 means the font is only available to the current process.
        # We also need to keep track of loaded fonts if we wanted to remove them, 
        # but for this app we just load them for the session.
        success = ctypes.windll.gdi32.AddFontResourceExW(font_path, 0x10, 0)
        return success > 0
    except Exception as e:
        print(f"Error loading font {font_path}: {e}")
        return False

class FontInfo:
    def __init__(self, family, path):
        self.family = family
        self.path = path
        self.directory = os.path.dirname(path)
        self.is_favorite = False

    def __repr__(self):
        return f"FontInfo({self.family}, {self.path})"

class DataManager:
    def __init__(self, filename="config.json"):
        self.filename = filename
        self.data = {
            "custom_dirs": [],
            "favorites": [],
            "projects": {}  # name -> { "lists": { "list_name": [font_family, ...] } }
        }
        self.load()

    def load(self):
        if os.path.exists(self.filename):
            try:
                with open(self.filename, 'r') as f:
                    loaded_data = json.load(f)
                    # Safely merge loaded data with defaults to ensure all keys exist
                    for key in self.data:
                        if key in loaded_data:
                            self.data[key] = loaded_data[key]
            except Exception:
                pass

    def save(self):
        try:
            with open(self.filename, 'w') as f:
                json.dump(self.data, f, indent=4)
        except Exception as e:
            print(f"Error saving data: {e}")

class ScrollableFontList(tk.Frame):
    def __init__(self, parent, colors, on_click, on_fav, app, show_add_btn=True):
        super().__init__(parent, bg=colors["bg"])
        self.colors = colors
        self.on_click = on_click
        self.on_fav = on_fav
        self.app = app
        self.show_add_btn = show_add_btn
        
        self.canvas = tk.Canvas(self, bg=colors["list_bg"], highlightthickness=0)
        self.scrollbar = tk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg=colors["list_bg"])
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        self.canvas.bind("<Configure>", self.on_canvas_configure)
        
        # Mouse wheel scrolling
        self.canvas.bind("<Enter>", self._bind_mousewheel)
        self.canvas.bind("<Leave>", self._unbind_mousewheel)
        
        # Keyboard jump
        self.canvas.bind("<KeyPress>", self._on_key_press)
        self.canvas.bind("<Button-1>", lambda e: self.canvas.focus_set())
        
        self.fonts_data = []

    def _on_key_press(self, event):
        if not event.char or not (event.char.isalnum() or event.char.isspace()):
            return
        
        char = event.char.lower()
        for i, f_info in enumerate(self.fonts_data):
            if f_info.family.lower().startswith(char):
                # Calculate scroll position
                # This is a bit tricky with canvas.yview_moveto
                # We can use the row height if it's constant, but it's not strictly constant here due to padding
                # Alternatively, we can use the relative position of the row in the frame
                children = self.scrollable_frame.winfo_children()
                if i < len(children):
                    target_widget = children[i]
                    self.scrollable_frame.update_idletasks() # Ensure layout is ready
                    y_pos = target_widget.winfo_y()
                    total_height = self.scrollable_frame.winfo_height()
                    if total_height > 0:
                        self.canvas.yview_moveto(y_pos / total_height)
                break

    def _bind_mousewheel(self, event):
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _unbind_mousewheel(self, event):
        self.canvas.unbind_all("<MouseWheel>")

    def on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def set_fonts(self, fonts):
        self.fonts_data = fonts
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
            
        for f_info in fonts:
            self.create_font_row(f_info)

    def create_font_row(self, f_info):
        row = tk.Frame(self.scrollable_frame, bg=self.colors["list_bg"], pady=5, padx=10)
        row.pack(fill=tk.X)
        
        if self.show_add_btn:
            add_btn = tk.Label(row, text="+", bg=self.colors["list_bg"], fg="#444444", font=("Arial", 16, "bold"), cursor="hand2")
            add_btn.pack(side=tk.LEFT, padx=5)
            add_btn.bind("<Button-1>", lambda e, f=f_info.family: [self.app.add_font_to_current_project_list(f), self.canvas.focus_set()])
        
        # Font Name Label
        lbl_font = (f_info.family, 30)
        
        name_lbl = tk.Label(row, text=f_info.family, bg=self.colors["list_bg"], fg=self.colors["text"], font=lbl_font, cursor="hand2")
        name_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)
        name_lbl.bind("<Button-1>", lambda e, f=f_info.family: [self.on_click(f), self.canvas.focus_set()])
        
        # Fav Button
        star = "★" if f_info.is_favorite else "☆"
        fav_btn = tk.Label(row, text=star, bg=self.colors["list_bg"], fg="gold", font=("Arial", 16), cursor="hand2")
        fav_btn.pack(side=tk.RIGHT, padx=5)
        fav_btn.bind("<Button-1>", lambda e, f=f_info.family: [self.on_fav(f), self.canvas.focus_set()])

    def show_add_to_list_menu(self, event, family):
        menu = tk.Menu(self, tearoff=0)
        projects = self.app.data_manager.data["projects"]
        if not projects:
            menu.add_command(label="No projects created", state="disabled")
        else:
            for p_name, p_data in projects.items():
                submenu = tk.Menu(menu, tearoff=0)
                for l_name in p_data["lists"]:
                    submenu.add_command(label=l_name, command=lambda p=p_name, l=l_name, f=family: self.app.add_font_to_project_list(p, l, f))
                menu.add_cascade(label=p_name, menu=submenu)
        menu.post(event.x_root, event.y_root)

class KorianFontsManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Korian Fonts Manager")
        self.root.geometry("1200x800")
        
        self.data_manager = DataManager()
        self.all_fonts = {} # family -> FontInfo
        self.fonts_lock = threading.RLock()
        self.last_clicked_font = None
        self.current_project = None
        self.sort_by_dir = False
        
        # Preview Settings
        self.preview_font_size = 72
        self.preview_fg = "black"
        self.preview_bg = "#CCCCCC"
        self.bg_image = None
        self.bg_image_tk = None
        self.bg_x = 0
        self.bg_y = 0
        self.bg_zoom = 1.0
        self.drag_start_x = 0
        self.drag_start_y = 0
        
        self.show_bounding_box = tk.BooleanVar(value=False)
        self.word_wrap = tk.BooleanVar(value=True)
        self.text_align = tk.StringVar(value="left")
        
        # Bounding Box (fractions of canvas size)
        self.bbox_left = 0.1
        self.bbox_top = 0.1
        self.bbox_right = 0.9
        self.bbox_bottom = 0.9
        
        self.dragging_bbox_move = False
        self.dragging_bbox_edge = None
        self.full_width_preview = False
        self._wheel_timer = None
        self._cached_bg_image = None
        self._cached_bg_zoom = 0
        
        # History for Undo/Redo
        self.undo_stack = []
        self.redo_stack = []
        self.ignore_history = True
        self.bg_image_path = None
        
        # UI Theme
        self.colors = {
            "bg": "#444444",         # Medium gray
            "sidebar_bg": "#333333", # Darker gray
            "list_bg": "#CCCCCC",    # Lightest gray for font areas
            "text": "black",         # Black text for font display
            "sidebar_text": "white", # White text for sidebar labels
            "active_tab": "#555555"
        }
        
        self.root.configure(bg=self.colors["bg"])
        self.setup_ui()
        self.ignore_history = False
        self.root.update()
        
        # Initial font scan
        threading.Thread(target=self.initial_font_scan, daemon=True).start()

    def setup_ui(self):
        # Bind virtual event for font updates
        self.root.bind("<<UpdateFonts>>", lambda e: self.update_font_lists())

        # Bind Undo/Redo keys
        self.root.bind("<Control-z>", lambda e: self.undo())
        self.root.bind("<Control-Z>", lambda e: self.undo())
        self.root.bind("<Control-Shift-Z>", lambda e: self.redo())
        self.root.bind("<Control-Shift-z>", lambda e: self.redo())

        # Vertical Paned Window to separate main content from preview
        self.vertical_paned = tk.PanedWindow(self.root, orient=tk.VERTICAL, bg=self.colors["bg"], borderwidth=0, sashwidth=4)
        self.vertical_paned.pack(fill=tk.BOTH, expand=True)

        # Horizontal Paned Window for 3 columns
        self.main_paned = tk.PanedWindow(self.vertical_paned, orient=tk.HORIZONTAL, bg=self.colors["bg"], borderwidth=0, sashwidth=4)
        self.vertical_paned.add(self.main_paned, stretch="always")

        # Left Column: Projects
        self.left_frame = tk.Frame(self.main_paned, bg=self.colors["sidebar_bg"], width=200)
        self.main_paned.add(self.left_frame, stretch="never")
        self.setup_left_column()

        # Center Column: Project Lists
        self.center_frame = tk.Frame(self.main_paned, bg=self.colors["bg"], width=400)
        self.main_paned.add(self.center_frame, stretch="always")
        self.setup_center_column()

        # Right Column: All Fonts / Favorites
        self.right_frame = tk.Frame(self.main_paned, bg=self.colors["bg"], width=400)
        self.main_paned.add(self.right_frame, stretch="always")
        self.setup_right_column()

        # Bottom Row: Custom Text Preview
        self.bottom_frame = tk.Frame(self.vertical_paned, bg=self.colors["sidebar_bg"])
        self.vertical_paned.add(self.bottom_frame, stretch="never", minsize=150)
        self.setup_bottom_row()

    def setup_left_column(self):
        tk.Label(self.left_frame, text="PROJECTS", bg=self.colors["sidebar_bg"], fg="white", font=("Arial", 12, "bold")).pack(pady=10)
        
        self.projects_listbox = tk.Listbox(self.left_frame, bg=self.colors["sidebar_bg"], fg="white", 
                                          selectbackground=self.colors["active_tab"], borderwidth=0, 
                                          highlightthickness=0, font=("Arial", 10))
        self.projects_listbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.projects_listbox.bind("<<ListboxSelect>>", self.on_project_select)
        self.projects_listbox.bind("<Button-3>", self.on_project_right_click)
        
        btn_frame = tk.Frame(self.left_frame, bg=self.colors["sidebar_bg"])
        btn_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=10)
        
        tk.Button(btn_frame, text="Add Project", command=self.add_project).pack(fill=tk.X, padx=10, pady=2)
        tk.Button(btn_frame, text="Add Font Dir", command=self.add_font_dir).pack(fill=tk.X, padx=10, pady=2)
        
        # Populate projects
        for project_name in self.data_manager.data["projects"]:
            self.projects_listbox.insert(tk.END, project_name)

    def setup_center_column(self):
        header_frame = tk.Frame(self.center_frame, bg=self.colors["bg"])
        header_frame.pack(fill=tk.X)
        tk.Label(header_frame, text="PROJECT LISTS", bg=self.colors["bg"], fg="white", font=("Arial", 12, "bold")).pack(side=tk.LEFT, padx=10, pady=10)
        tk.Button(header_frame, text="+ List", command=self.add_list, bg=self.colors["sidebar_bg"], fg="white").pack(side=tk.RIGHT, padx=10)
        
        self.center_notebook = ttk.Notebook(self.center_frame)
        self.center_notebook.pack(fill=tk.BOTH, expand=True)
        self.center_notebook.bind("<Button-3>", self.on_tab_right_click)
        
        # Style for Notebook
        style = ttk.Style()
        style.theme_use('default')
        style.configure('TNotebook', background=self.colors["bg"], borderwidth=0)
        style.configure('TNotebook.Tab', background=self.colors["sidebar_bg"], foreground="white", padding=[10, 5])
        style.map('TNotebook.Tab', background=[('selected', self.colors["active_tab"])])
        
        self.center_notebook.bind("<<NotebookTabChanged>>", self.on_center_tab_change)

    def setup_right_column(self):
        # Sort button
        sort_btn = tk.Button(self.right_frame, text="Sort: A-Z", command=self.toggle_sort)
        sort_btn.pack(side=tk.TOP, fill=tk.X)
        self.sort_btn = sort_btn

        self.right_notebook = ttk.Notebook(self.right_frame)
        self.right_notebook.pack(fill=tk.BOTH, expand=True)
        self.right_notebook.bind("<<NotebookTabChanged>>", self.on_right_tab_change)
        
        # All Fonts Tab
        self.all_fonts_frame = tk.Frame(self.right_notebook, bg=self.colors["bg"])
        self.right_notebook.add(self.all_fonts_frame, text="All Fonts")
        
        self.all_fonts_list = ScrollableFontList(self.all_fonts_frame, self.colors, self.on_font_click, self.toggle_favorite, self)
        self.all_fonts_list.pack(fill=tk.BOTH, expand=True)
        
        # Favorites Tab
        self.fav_fonts_frame = tk.Frame(self.right_notebook, bg=self.colors["bg"])
        self.right_notebook.add(self.fav_fonts_frame, text="Favorites")
        
        self.fav_fonts_list = ScrollableFontList(self.fav_fonts_frame, self.colors, self.on_font_click, self.toggle_favorite, self)
        self.fav_fonts_list.pack(fill=tk.BOTH, expand=True)

    def setup_bottom_row(self):
        # Tools row for colors, image upload, and bounding box toggle
        tools_frame = tk.Frame(self.bottom_frame, bg=self.colors["sidebar_bg"])
        tools_frame.pack(fill=tk.X, side=tk.TOP, padx=10, pady=2)
        
        tk.Label(tools_frame, text="Preview Text:", bg=self.colors["sidebar_bg"], fg="white").pack(side=tk.LEFT)
        self.preview_text_box = tk.Text(tools_frame, width=30, height=2)
        self.preview_text_box.pack(side=tk.LEFT, padx=5)
        self.preview_text_box.insert("1.0", "The quick brown fox jumps over the lazy dog")
        self.preview_text_box.bind("<FocusIn>", lambda e: self.save_to_history())
        self.preview_text_box.bind("<KeyRelease>", lambda e: self.update_preview())
        
        # Color Buttons
        tk.Button(tools_frame, text="Font Color", command=self.choose_fg_color, bg=self.colors["bg"], fg="white").pack(side=tk.LEFT, padx=2)
        self.fg_hex_entry = tk.Entry(tools_frame, width=8, bg=self.colors["bg"], fg="white", insertbackground="white")
        self.fg_hex_entry.pack(side=tk.LEFT, padx=2)
        self.fg_hex_entry.insert(0, self.preview_fg)
        self.fg_hex_entry.bind("<Return>", self.on_fg_hex_change)
        self.fg_hex_entry.bind("<FocusOut>", self.on_fg_hex_change)

        tk.Button(tools_frame, text="BG Color", command=self.choose_bg_color, bg=self.colors["bg"], fg="white").pack(side=tk.LEFT, padx=2)
        self.bg_hex_entry = tk.Entry(tools_frame, width=8, bg=self.colors["bg"], fg="white", insertbackground="white")
        self.bg_hex_entry.pack(side=tk.LEFT, padx=2)
        self.bg_hex_entry.insert(0, self.preview_bg)
        self.bg_hex_entry.bind("<Return>", self.on_bg_hex_change)
        self.bg_hex_entry.bind("<FocusOut>", self.on_bg_hex_change)
        
        tk.Button(tools_frame, text="Upload BG", command=self.upload_bg_image, bg=self.colors["bg"], fg="white").pack(side=tk.LEFT, padx=2)

        # Bounding Box Checkbox
        tk.Checkbutton(tools_frame, text="Bounding Box", variable=self.show_bounding_box, command=self.update_preview, 
                       bg=self.colors["sidebar_bg"], fg="white", selectcolor=self.colors["bg"]).pack(side=tk.LEFT, padx=5)
        
        # Word Wrap and Alignment
        tk.Checkbutton(tools_frame, text="Wrap", variable=self.word_wrap, command=self.update_preview, 
                       bg=self.colors["sidebar_bg"], fg="white", selectcolor=self.colors["bg"]).pack(side=tk.LEFT, padx=5)
        
        tk.Label(tools_frame, text="Align:", bg=self.colors["sidebar_bg"], fg="white").pack(side=tk.LEFT)
        align_combo = ttk.Combobox(tools_frame, textvariable=self.text_align, values=["left", "center", "right"], width=7)
        align_combo.pack(side=tk.LEFT, padx=2)
        align_combo.bind("<<ComboboxSelected>>", lambda e: self.update_preview())

        # Preview Canvas
        self.preview_canvas = tk.Canvas(self.bottom_frame, bg=self.preview_bg, highlightthickness=0)
        self.preview_canvas.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.preview_canvas.bind("<MouseWheel>", self.on_canvas_wheel)
        self.preview_canvas.bind("<Button-1>", self.on_canvas_drag_start)
        self.preview_canvas.bind("<B1-Motion>", self.on_canvas_drag_motion)
        self.preview_canvas.bind("<ButtonRelease-1>", self.on_canvas_drag_release)
        self.preview_canvas.bind("<Motion>", self.on_canvas_mouse_move)
        self.preview_canvas.bind("<Configure>", lambda e: self.update_preview())

    def get_state_snapshot(self):
        # Data state
        data_copy = json.loads(json.dumps(self.data_manager.data))
        
        # Preview state
        preview_state = {
            "last_clicked_font": self.last_clicked_font,
            "current_project": self.current_project,
            "preview_font_size": self.preview_font_size,
            "preview_fg": self.preview_fg,
            "preview_bg": self.preview_bg,
            "bg_x": self.bg_x,
            "bg_y": self.bg_y,
            "bg_zoom": self.bg_zoom,
            "bg_image_path": self.bg_image_path,
            "bbox_left": self.bbox_left,
            "bbox_top": self.bbox_top,
            "bbox_right": self.bbox_right,
            "bbox_bottom": self.bbox_bottom,
            "show_bounding_box": self.show_bounding_box.get(),
            "word_wrap": self.word_wrap.get(),
            "text_align": self.text_align.get(),
            "preview_text": self.preview_text_box.get("1.0", tk.END),
            "full_width_preview": self.full_width_preview,
            "sort_by_dir": self.sort_by_dir,
            "center_tab_idx": self.center_notebook.index("current") if self.center_notebook.tabs() else 0,
            "right_tab_idx": self.right_notebook.index("current") if self.right_notebook.tabs() else 0
        }
        return {"data": data_copy, "preview": preview_state}

    def load_state_snapshot(self, snapshot):
        self.data_manager.data = snapshot["data"]
        self.data_manager.save()
        
        # Synchronize favorites status in FontInfo objects
        with self.fonts_lock:
            favs = self.data_manager.data["favorites"]
            for family, f_info in self.all_fonts.items():
                f_info.is_favorite = (family in favs)
        
        p = snapshot["preview"]
        self.last_clicked_font = p["last_clicked_font"]
        self.current_project = p["current_project"]
        self.preview_font_size = p["preview_font_size"]
        self.preview_fg = p["preview_fg"]
        self.preview_bg = p["preview_bg"]
        self.bg_x = p["bg_x"]
        self.bg_y = p["bg_y"]
        self.bg_zoom = p["bg_zoom"]
        
        # Reload background image if path changed
        if self.bg_image_path != p["bg_image_path"]:
            self.bg_image_path = p["bg_image_path"]
            if self.bg_image_path:
                try:
                    self.bg_image = Image.open(self.bg_image_path)
                except Exception:
                    self.bg_image = None
            else:
                self.bg_image = None
        
        self.bbox_left = p["bbox_left"]
        self.bbox_top = p["bbox_top"]
        self.bbox_right = p["bbox_right"]
        self.bbox_bottom = p["bbox_bottom"]
        self.show_bounding_box.set(p["show_bounding_box"])
        self.word_wrap.set(p["word_wrap"])
        self.text_align.set(p["text_align"])
        
        self.preview_text_box.delete("1.0", tk.END)
        self.preview_text_box.insert("1.0", p["preview_text"].strip())
        
        self.full_width_preview = p["full_width_preview"]
        self.sort_by_dir = p.get("sort_by_dir", False)
        self.sort_btn.config(text="Sort: Dir + A-Z" if self.sort_by_dir else "Sort: A-Z")
        
        # Synchronize Hex Entries
        self.fg_hex_entry.delete(0, tk.END)
        self.fg_hex_entry.insert(0, self.preview_fg)
        self.bg_hex_entry.delete(0, tk.END)
        self.bg_hex_entry.insert(0, self.preview_bg)
        self.preview_canvas.config(bg=self.preview_bg)
        
        # Update UI components
        self.update_font_lists()
        self.update_project_listbox()
        self.refresh_project_lists()
        
        # Restore tab selections
        if "center_tab_idx" in p and p["center_tab_idx"] < len(self.center_notebook.tabs()):
            self.center_notebook.select(p["center_tab_idx"])
        if "right_tab_idx" in p and p["right_tab_idx"] < len(self.right_notebook.tabs()):
            self.right_notebook.select(p["right_tab_idx"])
            
        self.update_preview()
        self.arrange_layout()

    def save_to_history(self):
        if self.ignore_history:
            return
        self.undo_stack.append(self.get_state_snapshot())
        if len(self.undo_stack) > 100:
            self.undo_stack.pop(0)
        self.redo_stack.clear()

    def undo(self):
        if not self.undo_stack:
            return
        self.ignore_history = True
        try:
            self.redo_stack.append(self.get_state_snapshot())
            previous_state = self.undo_stack.pop()
            self.load_state_snapshot(previous_state)
            self.root.update() # Process pending events
        finally:
            self.ignore_history = False

    def redo(self):
        if not self.redo_stack:
            return
        self.ignore_history = True
        try:
            self.undo_stack.append(self.get_state_snapshot())
            next_state = self.redo_stack.pop()
            self.load_state_snapshot(next_state)
            self.root.update() # Process pending events
        finally:
            self.ignore_history = False

    def update_project_listbox(self):
        self.projects_listbox.delete(0, tk.END)
        for i, project_name in enumerate(self.data_manager.data["projects"]):
            self.projects_listbox.insert(tk.END, project_name)
            if project_name == self.current_project:
                self.projects_listbox.selection_clear(0, tk.END)
                self.projects_listbox.selection_set(i)

    def refresh_project_lists(self):
        # Save current selected tab index
        selected_idx = self.center_notebook.index("current") if self.center_notebook.tabs() else 0
        
        # Avoid saving history during refresh
        old_ignore = self.ignore_history
        self.ignore_history = True
        
        try:
            # Clear existing tabs and DESTROY their widgets to prevent memory leaks
            for tab in self.center_notebook.tabs():
                widget = self.center_notebook.nametowidget(tab)
                self.center_notebook.forget(tab)
                widget.destroy()
            
            if not self.current_project:
                return
                
            p_data = self.data_manager.data["projects"].get(self.current_project)
            if not p_data:
                return
                
            # Ensure Default list exists
            if "Default" not in p_data["lists"]:
                p_data["lists"]["Default"] = []
                self.data_manager.save()
                
            for list_name, families in p_data["lists"].items():
                frame = tk.Frame(self.center_notebook, bg=self.colors["bg"])
                self.center_notebook.add(frame, text=list_name)
                
                f_list = ScrollableFontList(frame, self.colors, self.on_font_click, self.toggle_favorite, self, show_add_btn=False)
                f_list.pack(fill=tk.BOTH, expand=True)
                
                list_fonts = []
                with self.fonts_lock:
                    for family in families:
                        if family in self.all_fonts:
                            list_fonts.append(self.all_fonts[family])
                
                f_list.set_fonts(sorted(list_fonts, key=lambda x: x.family.lower()))
                
            # Restore tab selection
            if selected_idx < len(self.center_notebook.tabs()):
                self.center_notebook.select(selected_idx)
            
            self.root.update() # Process events triggered by Notebook manipulation
        finally:
            self.ignore_history = old_ignore

    def on_project_right_click(self, event):
        if self.projects_listbox.size() == 0: return
        index = self.projects_listbox.nearest(event.y)
        if index < 0: return
        project_name = self.projects_listbox.get(index)
        if not project_name: return
        
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label=f"Delete Project '{project_name}'", command=lambda: self.delete_project(project_name))
        menu.post(event.x_root, event.y_root)

    def on_tab_right_click(self, event):
        try:
            index = self.center_notebook.index(f"@{event.x},{event.y}")
            list_name = self.center_notebook.tab(index, "text")
            
            menu = tk.Menu(self.root, tearoff=0)
            menu.add_command(label=f"Delete List '{list_name}'", command=lambda: self.delete_list(list_name))
            menu.post(event.x_root, event.y_root)
        except Exception:
            pass

    def delete_project(self, name):
        if messagebox.askyesno("Delete Project", f"Are you sure you want to delete project '{name}'?"):
            self.save_to_history()
            if name in self.data_manager.data["projects"]:
                del self.data_manager.data["projects"][name]
                if self.current_project == name:
                    self.current_project = None
                self.data_manager.save()
                self.update_project_listbox()
                self.refresh_project_lists()

    def add_list(self):
        if not self.current_project:
            messagebox.showinfo("Add List", "Please select a project first.")
            return
        
        name = simpledialog.askstring("Add List", "List Name:")
        if name:
            p_data = self.data_manager.data["projects"][self.current_project]
            if name in p_data["lists"]:
                messagebox.showerror("Error", "List already exists.")
                return
            self.save_to_history()
            p_data["lists"][name] = []
            self.data_manager.save()
            self.refresh_project_lists()
            # Select the new list
            self.center_notebook.select(len(self.center_notebook.tabs()) - 1)

    def delete_list(self, list_name):
        if not self.current_project:
            return
        
        if messagebox.askyesno("Delete List", f"Are you sure you want to delete list '{list_name}' from project '{self.current_project}'?"):
            self.save_to_history()
            p_data = self.data_manager.data["projects"][self.current_project]
            if list_name in p_data["lists"]:
                del p_data["lists"][list_name]
                self.data_manager.save()
                self.refresh_project_lists()

    def choose_fg_color(self):
        color = colorchooser.askcolor(initialcolor=self.preview_fg)[1]
        if color:
            self.save_to_history()
            self.preview_fg = color
            self.fg_hex_entry.delete(0, tk.END)
            self.fg_hex_entry.insert(0, color)
            self.update_preview()

    def on_fg_hex_change(self, event=None):
        color = self.fg_hex_entry.get()
        if color.startswith("#") and (len(color) == 7 or len(color) == 4):
            if color != self.preview_fg:
                self.save_to_history()
                self.preview_fg = color
                self.update_preview()

    def choose_bg_color(self):
        color = colorchooser.askcolor(initialcolor=self.preview_bg)[1]
        if color:
            self.save_to_history()
            self.preview_bg = color
            self.bg_hex_entry.delete(0, tk.END)
            self.bg_hex_entry.insert(0, color)
            self.preview_canvas.config(bg=color)
            self.update_preview()

    def on_bg_hex_change(self, event=None):
        color = self.bg_hex_entry.get()
        if color.startswith("#") and (len(color) == 7 or len(color) == 4):
            if color != self.preview_bg:
                self.save_to_history()
                self.preview_bg = color
                self.preview_canvas.config(bg=color)
                self.update_preview()

    def upload_bg_image(self):
        path = filedialog.askopenfilename(filetypes=[("Image files", "*.jpg *.jpeg *.png *.bmp")])
        if path:
            self.save_to_history()
            self.bg_image_path = path
            self.bg_image = Image.open(path)
            self._cached_bg_image = None # Invalidate cache
            self.bg_x = 0
            self.bg_y = 0
            self.bg_zoom = 1.0
            self.update_preview()

    def on_canvas_wheel(self, event):
        # Control key state for zooming image, otherwise zooming font
        if event.state & 0x0004: # Control key
            if self.bg_image:
                if event.delta > 0: self.bg_zoom *= 1.1
                else: self.bg_zoom /= 1.1
                self.update_preview()
        else:
            if event.delta > 0: self.preview_font_size += 2
            else: self.preview_font_size = max(4, self.preview_font_size - 2)
            self.update_preview()
        
        # Debounce history save
        if self._wheel_timer:
            self.root.after_cancel(self._wheel_timer)
        self._wheel_timer = self.root.after(500, self.save_to_history)
        
        return "break" # Prevent propagation to scrollable lists

    def on_canvas_drag_start(self, event):
        self.drag_start_x = event.x
        self.drag_start_y = event.y
        self.dragging_bbox_edge = None
        self.dragging_bbox_move = False
        
        if self.show_bounding_box.get():
            cw = self.preview_canvas.winfo_width()
            ch = self.preview_canvas.winfo_height()
            bx1, by1 = self.bbox_left * cw, self.bbox_top * ch
            bx2, by2 = self.bbox_right * cw, self.bbox_bottom * ch
            
            handle_size = 10
            # Check move handle (top-right)
            if bx2 - handle_size < event.x < bx2 and by1 < event.y < by1 + handle_size:
                self.dragging_bbox_move = True
                return

            tolerance = 5
            # Priority to edges
            if abs(event.x - bx1) < tolerance and by1 - tolerance < event.y < by2 + tolerance:
                self.dragging_bbox_edge = "left"
            elif abs(event.x - bx2) < tolerance and by1 - tolerance < event.y < by2 + tolerance:
                self.dragging_bbox_edge = "right"
            elif abs(event.y - by1) < tolerance and bx1 - tolerance < event.x < bx2 + tolerance:
                self.dragging_bbox_edge = "top"
            elif abs(event.y - by2) < tolerance and bx1 - tolerance < event.x < bx2 + tolerance:
                self.dragging_bbox_edge = "bottom"

    def on_canvas_drag_motion(self, event):
        cw = self.preview_canvas.winfo_width()
        ch = self.preview_canvas.winfo_height()
        
        if self.dragging_bbox_edge:
            if self.dragging_bbox_edge == "left":
                self.bbox_left = max(0, min(event.x / cw, self.bbox_right - 0.01))
            elif self.dragging_bbox_edge == "right":
                self.bbox_right = min(1, max(event.x / cw, self.bbox_left + 0.01))
            elif self.dragging_bbox_edge == "top":
                self.bbox_top = max(0, min(event.y / ch, self.bbox_bottom - 0.01))
            elif self.dragging_bbox_edge == "bottom":
                self.bbox_bottom = min(1, max(event.y / ch, self.bbox_top + 0.01))
            self.update_preview()
        elif getattr(self, "dragging_bbox_move", False):
            dx = (event.x - self.drag_start_x) / cw
            dy = (event.y - self.drag_start_y) / ch
            
            # Move whole box while keeping in bounds (with clamping)
            if self.bbox_left + dx < 0: dx = -self.bbox_left
            if self.bbox_right + dx > 1: dx = 1 - self.bbox_right
            if self.bbox_top + dy < 0: dy = -self.bbox_top
            if self.bbox_bottom + dy > 1: dy = 1 - self.bbox_bottom
            
            self.bbox_left += dx
            self.bbox_right += dx
            self.bbox_top += dy
            self.bbox_bottom += dy
                
            self.drag_start_x = event.x
            self.drag_start_y = event.y
            self.update_preview()
        elif self.bg_image:
            dx = event.x - self.drag_start_x
            dy = event.y - self.drag_start_y
            self.bg_x += dx
            self.bg_y += dy
            self.drag_start_x = event.x
            self.drag_start_y = event.y
            self.update_preview()

    def on_canvas_drag_release(self, event):
        if self.dragging_bbox_edge or getattr(self, "dragging_bbox_move", False):
            self.save_to_history()
        elif self.bg_image:
            self.save_to_history()
        self.dragging_bbox_edge = None
        self.dragging_bbox_move = False

    def on_canvas_mouse_move(self, event):
        if not self.show_bounding_box.get():
            self.preview_canvas.config(cursor="")
            return
            
        cw = self.preview_canvas.winfo_width()
        ch = self.preview_canvas.winfo_height()
        bx1, by1 = self.bbox_left * cw, self.bbox_top * ch
        bx2, by2 = self.bbox_right * cw, self.bbox_bottom * ch
        
        handle_size = 10
        if bx2 - handle_size < event.x < bx2 and by1 < event.y < by1 + handle_size:
            self.preview_canvas.config(cursor="fleur")
            return

        tolerance = 5
        if abs(event.x - bx1) < tolerance and by1 - tolerance < event.y < by2 + tolerance:
            self.preview_canvas.config(cursor="size_we")
        elif abs(event.x - bx2) < tolerance and by1 - tolerance < event.y < by2 + tolerance:
            self.preview_canvas.config(cursor="size_we")
        elif abs(event.y - by1) < tolerance and bx1 - tolerance < event.x < bx2 + tolerance:
            self.preview_canvas.config(cursor="size_ns")
        elif abs(event.y - by2) < tolerance and bx1 - tolerance < event.x < bx2 + tolerance:
            self.preview_canvas.config(cursor="size_ns")
        else:
            self.preview_canvas.config(cursor="")

    def initial_font_scan(self):
        dirs = [os.path.join(os.environ['WINDIR'], 'Fonts')]
        user_font_dir = os.path.join(os.environ['LOCALAPPDATA'], 'Microsoft', 'Windows', 'Fonts')
        if os.path.exists(user_font_dir):
            dirs.append(user_font_dir)
        dirs.extend(self.data_manager.data["custom_dirs"])
        
        self.scan_directories(dirs)
        self.root.event_generate("<<UpdateFonts>>", when="tail")

    def scan_directories(self, directories):
        # Copy favorites to avoid shared access issues in thread
        favorites_copy = list(self.data_manager.data["favorites"])
        for directory in directories:
            if not os.path.exists(directory):
                continue
            for root, _, files in os.walk(directory):
                for file in files:
                    if file.lower().endswith(('.ttf', '.otf')):
                        path = os.path.join(root, file)
                        try:
                            f = ImageFont.truetype(path)
                            family = f.getname()[0]
                            with self.fonts_lock:
                                if family not in self.all_fonts:
                                    load_font(path)
                                    self.all_fonts[family] = FontInfo(family, path)
                                    if family in favorites_copy:
                                        self.all_fonts[family].is_favorite = True
                        except:
                            continue

    def toggle_sort(self):
        self.save_to_history()
        self.sort_by_dir = not self.sort_by_dir
        self.sort_btn.config(text="Sort: Dir + A-Z" if self.sort_by_dir else "Sort: A-Z")
        self.update_font_lists()

    def update_font_lists(self):
        with self.fonts_lock:
            all_fonts_list = list(self.all_fonts.values())
            
        if self.sort_by_dir:
            sorted_fonts = sorted(all_fonts_list, key=lambda x: (x.directory.lower(), x.family.lower()))
        else:
            sorted_fonts = sorted(all_fonts_list, key=lambda x: x.family.lower())
        
        self.all_fonts_list.set_fonts(sorted_fonts)
        
        fav_fonts = [f for f in sorted_fonts if f.is_favorite]
        self.fav_fonts_list.set_fonts(fav_fonts)

    def add_font_to_current_project_list(self, family):
        list_name = self.get_current_list_name()
        if self.current_project and list_name:
            self.add_font_to_project_list(self.current_project, list_name, family)
        else:
            messagebox.showinfo("Add Font", "Please select a project list first.")

    def get_current_list_name(self):
        try:
            current_tab = self.center_notebook.select()
            if not current_tab: return None
            tab_text = self.center_notebook.tab(current_tab, "text")
            if tab_text == "+": return None
            return tab_text
        except:
            return None

    def on_font_click(self, family):
        self.save_to_history()
        self.last_clicked_font = family
        self.update_preview()

    def toggle_favorite(self, family):
        with self.fonts_lock:
            if family in self.all_fonts:
                self.save_to_history()
                self.all_fonts[family].is_favorite = not self.all_fonts[family].is_favorite
                if self.all_fonts[family].is_favorite:
                    if family not in self.data_manager.data["favorites"]:
                        self.data_manager.data["favorites"].append(family)
                else:
                    if family in self.data_manager.data["favorites"]:
                        self.data_manager.data["favorites"].remove(family)
                self.data_manager.save()
        
        # Update UI outside lock
        self.update_font_lists()
        self.refresh_project_lists()

    def update_preview(self):
        if not hasattr(self, 'preview_canvas'): return
        
        # Cancel pending update if any
        if hasattr(self, '_preview_timer'):
            self.root.after_cancel(self._preview_timer)
            del self._preview_timer
            
        self.preview_canvas.delete("all")
        cw = self.preview_canvas.winfo_width()
        ch = self.preview_canvas.winfo_height()
        if cw < 2 or ch < 2: 
            # If not yet laid out, schedule update
            self._preview_timer = self.root.after(100, self.update_preview)
            return

        # 1. Background Image
        if self.bg_image:
            if self._cached_bg_zoom != self.bg_zoom or self._cached_bg_image is None:
                iw, ih = self.bg_image.size
                new_size = (int(iw * self.bg_zoom), int(ih * self.bg_zoom))
                if new_size[0] > 0 and new_size[1] > 0:
                    resized_img = self.bg_image.resize(new_size, Image.Resampling.LANCZOS)
                    self.bg_image_tk = ImageTk.PhotoImage(resized_img)
                    self._cached_bg_image = self.bg_image_tk
                    self._cached_bg_zoom = self.bg_zoom
                else:
                    self._cached_bg_image = None
            
            if self._cached_bg_image:
                self.preview_canvas.create_image(self.bg_x, self.bg_y, image=self._cached_bg_image, anchor="nw")

        # 2. Bounding Box
        bx1, by1 = self.bbox_left * cw, self.bbox_top * ch
        bx2, by2 = self.bbox_right * cw, self.bbox_bottom * ch
        
        if self.show_bounding_box.get():
            self.preview_canvas.create_rectangle(bx1, by1, bx2, by2, outline="black", dash=(4, 4))
            self.preview_canvas.create_rectangle(bx1, by1, bx2, by2, outline="white", dash=(4, 4), dashoffset=4)
            # Move handle in upper right
            handle_size = 10
            self.preview_canvas.create_rectangle(bx2 - handle_size, by1, bx2, by1 + handle_size, fill="white", outline="black")

        # 3. Text
        text = self.preview_text_box.get("1.0", "end-1c")
        if not text: text = "Preview Text"
        family = self.last_clicked_font if self.last_clicked_font else "Arial"
        
        font_config = (family, int(self.preview_font_size))
        
        # Word Wrap and Alignment
        wrap_width = 0
        if self.word_wrap.get():
            wrap_width = bx2 - bx1
        
        align = self.text_align.get()
        anchor = {"left": "nw", "center": "n", "right": "ne"}[align]
        tx = bx1 if align == "left" else (bx1 + bx2)/2 if align == "center" else bx2
        ty = by1
        
        try:
            self.preview_canvas.create_text(
                tx, ty, text=text, fill=self.preview_fg, font=font_config,
                anchor=anchor, width=wrap_width, justify=align
            )
        except:
            # Fallback to Arial if family fails
            self.preview_canvas.create_text(
                tx, ty, text=text, fill=self.preview_fg, font=("Arial", int(self.preview_font_size)),
                anchor=anchor, width=wrap_width, justify=align
            )

    def on_project_select(self, event):
        selection = self.projects_listbox.curselection()
        if selection:
            new_project = self.projects_listbox.get(selection[0])
            if new_project != self.current_project:
                self.save_to_history()
                self.current_project = new_project
                self.refresh_project_lists()

    def add_font_to_project_list(self, project_name, list_name, family):
        p_data = self.data_manager.data["projects"][project_name]
        if family not in p_data["lists"][list_name]:
            self.save_to_history()
            p_data["lists"][list_name].append(family)
            self.data_manager.save()
            if self.current_project == project_name:
                self.refresh_project_lists()

    def add_project(self):
        name = simpledialog.askstring("Add Project", "Project Name:")
        if name:
            if name in self.data_manager.data["projects"]:
                messagebox.showerror("Error", "Project already exists.")
                return
            self.save_to_history()
            self.data_manager.data["projects"][name] = {"lists": {"Default": []}}
            self.data_manager.save()
            self.update_project_listbox()

    def add_font_dir(self):
        directory = filedialog.askdirectory()
        if directory:
            if directory not in self.data_manager.data["custom_dirs"]:
                self.save_to_history()
                self.data_manager.data["custom_dirs"].append(directory)
                self.data_manager.save()
                threading.Thread(target=self.scan_and_update, args=([directory],), daemon=True).start()

    def scan_and_update(self, directories):
        self.scan_directories(directories)
        self.root.event_generate("<<UpdateFonts>>", when="tail")

    def on_center_tab_change(self, event):
        if not getattr(self, "ignore_history", False):
            # Only save if we have tabs (not during clear)
            if self.center_notebook.tabs():
                self.save_to_history()

    def on_right_tab_change(self, event):
        if not getattr(self, "ignore_history", False):
            if self.right_notebook.tabs():
                self.save_to_history()

if __name__ == "__main__":
    root = tk.Tk()
    app = KorianFontsManagerApp(root)
    root.mainloop()
