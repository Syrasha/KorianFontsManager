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
                    
                    # Migration: convert projects from {name: {lists: {name: [fonts]}}} to {name: [fonts]}
                    new_projects = {}
                    for p_name, p_data in self.data["projects"].items():
                        if isinstance(p_data, dict) and "lists" in p_data:
                            all_fonts_in_p = []
                            for l_fonts in p_data["lists"].values():
                                for f in l_fonts:
                                    if f not in all_fonts_in_p:
                                        all_fonts_in_p.append(f)
                            new_projects[p_name] = all_fonts_in_p
                        elif isinstance(p_data, list):
                            new_projects[p_name] = p_data
                        else:
                            new_projects[p_name] = []
                    self.data["projects"] = new_projects
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

    def reset_scroll(self):
        self.canvas.yview_moveto(0)

    def set_fonts(self, fonts):
        bg_color = self.app.preview_bg
        
        # Smart update: if families and order are same, just update existing widgets
        if self.fonts_data and [f.family for f in fonts] == [f.family for f in self.fonts_data]:
            rows = self.scrollable_frame.winfo_children()
            if len(rows) == len(fonts):
                for i, f_info in enumerate(fonts):
                    self.update_font_row(rows[i], f_info)
                
                # Update general backgrounds
                self.configure(bg=bg_color)
                self.canvas.configure(bg=bg_color)
                self.scrollable_frame.configure(bg=bg_color)
                self.fonts_data = fonts
                return

        self.fonts_data = fonts
        
        # Double buffering: build the new list in a hidden frame first to avoid flickering
        # We use a frame that is not yet placed as a window
        new_frame = tk.Frame(self.canvas, bg=bg_color)
        
        # Temporarily use new_frame for create_font_row
        old_frame = self.scrollable_frame
        self.scrollable_frame = new_frame
        
        for f_info in fonts:
            self.create_font_row(f_info)
            
        # Swap the frame in the canvas window and match current width
        self.canvas.itemconfig(self.canvas_window, window=new_frame, width=self.canvas.winfo_width())
        self.scrollable_frame = new_frame # Update permanent reference
        
        # Re-bind the configure event so the scroll region updates for the new frame
        new_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        # Update backgrounds
        self.configure(bg=bg_color)
        self.canvas.configure(bg=bg_color)
        
        # Clean up the old frame
        if old_frame and old_frame != new_frame:
            old_frame.destroy()

    def create_font_row(self, f_info):
        family = f_info.family
        bg_color = self.app.preview_bg
        
        row = tk.Frame(self.scrollable_frame, bg=bg_color, pady=5, padx=10)
        row.pack(fill=tk.X)
        
        # Enable dragging from the row components
        self.app.bind_font_drag(row, family)
        
        # Left side: Family name in small text
        info_frame = tk.Frame(row, bg=bg_color)
        info_frame.pack(side=tk.LEFT, padx=(0, 20))
        self.app.bind_font_drag(info_frame, family)

        name_lbl = tk.Label(info_frame, text=family, bg=bg_color, fg="#555555", font=("Arial", 8))
        name_lbl.pack(anchor="w")
        self.app.bind_font_drag(name_lbl, family)

        # Preview area using Canvas
        preview_canvas = tk.Canvas(row, bg=bg_color, highlightthickness=0, height=150)
        preview_canvas.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Initial draw
        self.setup_row_canvas(preview_canvas, f_info)

        # Fav Button
        star = "★" if f_info.is_favorite else "☆"
        fav_btn = tk.Label(row, text=star, bg=bg_color, fg="gold", font=("Arial", 16), cursor="hand2")
        fav_btn.pack(side=tk.RIGHT, padx=5)
        fav_btn.bind("<Button-1>", lambda e, f=family: [self.on_fav(f), self.canvas.focus_set()])

        # Store references for smart updates
        row.info_frame = info_frame
        row.name_lbl = name_lbl
        row.preview_canvas = preview_canvas
        row.fav_btn = fav_btn

    def setup_row_canvas(self, preview_canvas, f_info):
        family = f_info.family
        bg_color = self.app.preview_bg
        preview_text = self.app.preview_text_box.get("1.0", "end-1c") if hasattr(self.app, 'preview_text_box') else "Preview Text"
        if not preview_text.strip(): preview_text = "Preview Text"
        
        if family not in self.app.font_bboxes:
            self.app.font_bboxes[family] = (400, 200)
        bw, bh = self.app.font_bboxes[family]
        
        # Background image
        if self.app.bg_image_tk:
            preview_canvas.create_image(0, 0, image=self.app.bg_image_tk, anchor="nw", tags="bg_img")
            
        font_size = int(self.app.preview_font_size / 2)
        align = self.app.text_align.get() if hasattr(self.app, 'text_align') else "left"
        anchor = "nw"
        tx = 10
        if align == "center":
            anchor = "n"
            tx = bw / 2 + 5
        elif align == "right":
            anchor = "ne"
            tx = bw + 5

        # Draw items
        preview_canvas.create_text(tx, 10, text=preview_text, font=(family, font_size), 
                                   fill=self.app.preview_fg, anchor=anchor, justify=align, 
                                   width=bw if self.app.show_bounding_box.get() else 0, tags="text")
        
        secondary_font = (family, 12)
        alpha_text = "AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz"
        preview_canvas.create_text(tx, 50, text=alpha_text, font=secondary_font,
                                   fill=self.app.preview_fg, anchor=anchor, tags="secondary_alpha")
        
        punc_text = "1234567890 !?.,;:'\"()/\\{}[]-=_+~`@#$%^&*`\""
        preview_canvas.create_text(tx, 70, text=punc_text, font=secondary_font,
                                   fill=self.app.preview_fg, anchor=anchor, tags="secondary_punc")
        
        # Setup/Update bbox if enabled
        if self.app.show_bounding_box.get():
            self._draw_bbox(preview_canvas, family, bw, bh)

        self._update_canvas_layout(preview_canvas, family, tx, anchor, align, bw, bh)
        
        # Dragging and clicking
        preview_canvas.bind("<Button-1>", lambda e, f=family: self.on_click(f), add="+")
        self.app.bind_font_drag(preview_canvas, family)

    def update_font_row(self, row, f_info):
        family = f_info.family
        bg_color = self.app.preview_bg
        
        row.configure(bg=bg_color)
        row.info_frame.configure(bg=bg_color)
        row.name_lbl.configure(bg=bg_color)
        row.fav_btn.configure(bg=bg_color)
        
        star = "★" if f_info.is_favorite else "☆"
        row.fav_btn.configure(text=star)
        
        c = row.preview_canvas
        c.configure(bg=bg_color)
        
        # Update bg image
        c.delete("bg_img")
        if self.app.bg_image_tk:
            c.create_image(0, 0, image=self.app.bg_image_tk, anchor="nw", tags="bg_img")
            c.tag_lower("bg_img")
            
        preview_text = self.app.preview_text_box.get("1.0", "end-1c") if hasattr(self.app, 'preview_text_box') else "Preview Text"
        if not preview_text.strip(): preview_text = "Preview Text"
        
        bw, bh = self.app.font_bboxes.get(family, (400, 200))
        font_size = int(self.app.preview_font_size / 2)
        align = self.app.text_align.get()
        
        anchor = "nw"
        tx = 10
        if align == "center":
            anchor = "n"
            tx = bw / 2 + 5
        elif align == "right":
            anchor = "ne"
            tx = bw + 5

        # Update text items
        c.itemconfig("text", text=preview_text, font=(family, font_size), 
                     fill=self.app.preview_fg, anchor=anchor, justify=align, 
                     width=bw if self.app.show_bounding_box.get() else 0)
        
        secondary_font = (family, 12)
        c.itemconfig("secondary_alpha", font=secondary_font, fill=self.app.preview_fg, anchor=anchor)
        c.itemconfig("secondary_punc", font=secondary_font, fill=self.app.preview_fg, anchor=anchor)
        
        # Update bbox
        if self.app.show_bounding_box.get():
            if not c.find_withtag("bbox"):
                self._draw_bbox(c, family, bw, bh)
            else:
                c.coords("bbox", 5, 5, bw+5, bh+5)
                c.coords("handle", bw, bh, bw+10, bh+10)
        else:
            c.delete("bbox", "handle")

        self._update_canvas_layout(c, family, tx, anchor, align, bw, bh)

    def _draw_bbox(self, c, family, bw, bh):
        c.create_rectangle(5, 5, bw+5, bh+5, outline="red", dash=(4, 4), tags="bbox")
        c.create_rectangle(bw, bh, bw+10, bh+10, fill="red", tags="handle")
        
        c.tag_bind("handle", "<Button-1>", lambda e: setattr(c, '_drag_start', (e.x, e.y)))
        c.tag_bind("handle", "<B1-Motion>", lambda e, f=family, canvas=c: self._on_bbox_resize(e, f, canvas))
        c.tag_bind("handle", "<ButtonRelease-1>", lambda e: self.app.save_to_history())

    def _on_bbox_resize(self, e, f, c):
        ds = getattr(c, '_drag_start', (e.x, e.y))
        dx = e.x - ds[0]
        dy = e.y - ds[1]
        cur_w, cur_h = self.app.font_bboxes[f]
        new_w = max(20, cur_w + dx)
        new_h = max(20, cur_h + dy)
        self.app.font_bboxes[f] = (new_w, new_h)
        
        # Update UI directly
        align = self.app.text_align.get()
        anchor = "nw"
        tx = 10
        if align == "center":
            anchor = "n"
            tx = new_w / 2 + 5
        elif align == "right":
            anchor = "ne"
            tx = new_w + 5
            
        c.coords("bbox", 5, 5, new_w+5, new_h+5)
        c.coords("handle", new_w, new_h, new_w+10, new_h+10)
        c.itemconfig("text", width=new_w if self.app.show_bounding_box.get() else 0)
        self._update_canvas_layout(c, f, tx, anchor, align, new_w, new_h)
        c._drag_start = (e.x, e.y)

    def _update_canvas_layout(self, c, family, tx, anchor, align, bw, bh):
        c.coords("text", tx, 10)
        c.itemconfig("text", anchor=anchor, justify=align)
        
        text_bbox = c.bbox("text")
        next_y = text_bbox[3] + 10 if text_bbox else 50
        c.coords("secondary_alpha", tx, next_y)
        c.coords("secondary_punc", tx, next_y + 20)
        
        # Update canvas height
        p_bbox = c.bbox("secondary_punc")
        text_bottom = p_bbox[3] if p_bbox else 0
        h = text_bottom + 5
        if self.app.show_bounding_box.get():
            h = max(h, bh + 10)
        c.configure(height=max(h, 30))

    def show_add_to_list_menu(self, event, family):
        menu = tk.Menu(self, tearoff=0)
        projects = self.app.data_manager.data["projects"]
        if not projects:
            menu.add_command(label="No projects created", state="disabled")
        else:
            for p_name in sorted(projects.keys()):
                menu.add_command(label=p_name, command=lambda p=p_name, f=family: self.app.add_font_to_project(p, f))
        menu.post(event.x_root, event.y_root)

class ProjectsTree(tk.Frame):
    def __init__(self, parent, colors, app):
        super().__init__(parent, bg=colors["sidebar_bg"])
        self.colors = colors
        self.app = app
        self.expanded_projects = set()

        self.canvas = tk.Canvas(self, bg=colors["sidebar_bg"], highlightthickness=0)
        self.scrollbar = tk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg=colors["sidebar_bg"])

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(self.canvas_window, width=e.width))
        
        # Mouse wheel
        self.canvas.bind("<Enter>", lambda e: self.canvas.bind_all("<MouseWheel>", self._on_mousewheel))
        self.canvas.bind("<Leave>", lambda e: self.canvas.unbind_all("<MouseWheel>"))

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def update_tree(self):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        projects = self.app.data_manager.data["projects"]
        for p_name in sorted(projects.keys()):
            self.create_project_row(p_name)

    def create_project_row(self, p_name):
        p_frame = tk.Frame(self.scrollable_frame, bg=self.colors["sidebar_bg"])
        p_frame.pack(fill=tk.X)
        p_frame.project_name = p_name # Support drop on the entire project container

        is_expanded = p_name in self.expanded_projects
        symbol = "▼" if is_expanded else "▶"
        
        bg_color = self.colors["active_tab"] if p_name == self.app.current_project else self.colors["sidebar_bg"]

        row = tk.Frame(p_frame, bg=bg_color, pady=2)
        row.pack(fill=tk.X)
        row.project_name = p_name
        
        lbl = tk.Label(row, text=f"{symbol} {p_name}", bg=bg_color, fg="white", 
                       font=("Arial", 10, "bold"), anchor="w", cursor="hand2")
        lbl.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        lbl.project_name = p_name
        
        lbl.bind("<Button-1>", lambda e, n=p_name: self.toggle_project(n))
        lbl.bind("<Button-3>", lambda e, n=p_name: self.app.on_project_right_click(e, n))

        if is_expanded:
            fonts_frame = tk.Frame(p_frame, bg=self.colors["sidebar_bg"])
            fonts_frame.pack(fill=tk.X, padx=(20, 0))
            
            families = self.app.data_manager.data["projects"][p_name]
            for family in sorted(families):
                self.create_font_row(fonts_frame, p_name, family)

    def create_font_row(self, parent, p_name, family):
        f_row = tk.Frame(parent, bg=self.colors["sidebar_bg"], height=25)
        f_row.pack(fill=tk.X)
        f_row.pack_propagate(False)
        f_row.project_name = p_name  # Support as drop target
        
        # We need to check if font is loaded to show it in its typeface
        display_font = ("Arial", 10)
        with self.app.fonts_lock:
            if family in self.app.all_fonts:
                display_font = (family, 12)
        
        f_lbl = tk.Label(f_row, text=family, bg=self.colors["sidebar_bg"], fg="white",
                         font=display_font, anchor="w", cursor="hand2")
        f_lbl.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        f_lbl.project_name = p_name  # Support as drop target
        
        f_lbl.bind("<Button-1>", lambda e, f=family: self.app.on_font_click(f), add="+")
        self.app.bind_font_drag(f_lbl, family)
        
        f_lbl.bind("<Enter>", lambda e, l=f_lbl, f=family: self.on_font_hover(l, f), add="+")
        f_lbl.bind("<Leave>", lambda e, l=f_lbl, f=family: self.on_font_leave(l, f), add="+")
        f_lbl.bind("<Button-3>", lambda e, p=p_name, f=family: self.show_font_context_menu(e, p, f))

    def on_font_hover(self, label, family):
        label.configure(font=("Arial", 10))

    def on_font_leave(self, label, family):
        display_font = (family, 12)
        with self.app.fonts_lock:
            if family not in self.app.all_fonts:
                display_font = ("Arial", 10)
        label.configure(font=display_font)

    def toggle_project(self, p_name):
        if p_name in self.expanded_projects:
            self.expanded_projects.remove(p_name)
        else:
            self.expanded_projects.add(p_name)
        
        self.app.current_project = p_name
        
        # Filter the font list to show only this project's fonts
        families = self.app.data_manager.data["projects"].get(p_name, [])
        self.app.update_font_lists(families_filter=families)
        
        self.update_tree()

    def show_font_context_menu(self, event, p_name, family):
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label=f"Remove '{family}' from Project", 
                         command=lambda: self.app.remove_font_from_project(p_name, family))
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
        
        self.show_bounding_box = tk.BooleanVar(value=False)
        self.text_align = tk.StringVar(value="left")
        
        # Per-font bounding box: family -> (width, height)
        self.font_bboxes = {}
        
        self._size_timer = None
        
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

        # Horizontal Paned Window for columns
        self.main_paned = tk.PanedWindow(self.root, orient=tk.HORIZONTAL, bg=self.colors["bg"], borderwidth=0, sashwidth=4)
        self.main_paned.pack(fill=tk.BOTH, expand=True)

        # Left Column: Projects & Controls
        self.left_frame = tk.Frame(self.main_paned, bg=self.colors["sidebar_bg"], width=160)
        self.main_paned.add(self.left_frame, stretch="never", width=160)
        self.setup_left_column()

        # Right Column: Fonts with Repeating Preview
        self.right_frame = tk.Frame(self.main_paned, bg=self.colors["bg"], width=400)
        self.main_paned.add(self.right_frame, stretch="always")
        self.setup_right_column()

    def setup_left_column(self):
        tk.Label(self.left_frame, text="PROJECTS", bg=self.colors["sidebar_bg"], fg="white", font=("Arial", 12, "bold")).pack(pady=10)
        
        self.projects_tree = ProjectsTree(self.left_frame, self.colors, self)
        self.projects_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Add buttons for projects
        proj_btn_frame = tk.Frame(self.left_frame, bg=self.colors["sidebar_bg"])
        proj_btn_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Button(proj_btn_frame, text="Add Project", command=self.add_project).pack(fill=tk.X, pady=2)
        tk.Button(proj_btn_frame, text="Add Font Dir", command=self.add_font_dir).pack(fill=tk.X, pady=2)

        # Preview Controls area below projects
        tk.Label(self.left_frame, text="PREVIEW CONTROLS", bg=self.colors["sidebar_bg"], fg="white", font=("Arial", 10, "bold")).pack(pady=(10, 5))
        
        self.setup_preview_controls()

        # Populate projects
        self.projects_tree.update_tree()

    def setup_preview_controls(self):
        controls_frame = tk.Frame(self.left_frame, bg=self.colors["sidebar_bg"])
        controls_frame.pack(fill=tk.X, padx=10, pady=5)

        # Preview Text
        tk.Label(controls_frame, text="Preview Text:", bg=self.colors["sidebar_bg"], fg="white").pack(anchor="w")
        self.preview_text_box = tk.Text(controls_frame, width=10, height=2, wrap="word")
        self.preview_text_box.pack(fill=tk.X, pady=(0, 10))
        self.preview_text_box.insert("1.0", "Preview Text")
        self.preview_text_box.bind("<FocusIn>", lambda e: self.save_to_history())
        self.preview_text_box.bind("<KeyRelease>", lambda e: self.update_font_lists()) # Refresh all rows

        # Font Color Row
        fg_row = tk.Frame(controls_frame, bg=self.colors["sidebar_bg"])
        fg_row.pack(fill=tk.X, pady=2)
        
        tk.Button(fg_row, text="Font Color", command=self.choose_fg_color, bg=self.colors["bg"], fg="white", font=("Arial", 8)).pack(side=tk.LEFT, padx=2)
        self.fg_hex_entry = tk.Entry(fg_row, width=8, bg=self.colors["bg"], fg="white", insertbackground="white")
        self.fg_hex_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        self.fg_hex_entry.insert(0, self.preview_fg)
        self.fg_hex_entry.bind("<Return>", self.on_fg_hex_change)
        self.fg_hex_entry.bind("<FocusOut>", self.on_fg_hex_change)

        # BG Color Row
        bg_row = tk.Frame(controls_frame, bg=self.colors["sidebar_bg"])
        bg_row.pack(fill=tk.X, pady=2)

        tk.Button(bg_row, text="BG Color", command=self.choose_bg_color, bg=self.colors["bg"], fg="white", font=("Arial", 8)).pack(side=tk.LEFT, padx=2)
        self.bg_hex_entry = tk.Entry(bg_row, width=8, bg=self.colors["bg"], fg="white", insertbackground="white")
        self.bg_hex_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        self.bg_hex_entry.insert(0, self.preview_bg)
        self.bg_hex_entry.bind("<Return>", self.on_bg_hex_change)
        self.bg_hex_entry.bind("<FocusOut>", self.on_bg_hex_change)

        # Options Row (Checkboxes)
        options_row = tk.Frame(controls_frame, bg=self.colors["sidebar_bg"])
        options_row.pack(fill=tk.X, pady=2)
        
        tk.Checkbutton(options_row, text="Box", variable=self.show_bounding_box, command=self.update_font_lists, 
                       bg=self.colors["sidebar_bg"], fg="white", selectcolor=self.colors["bg"]).pack(side=tk.LEFT, padx=2)
        
        # Align Row
        align_row = tk.Frame(controls_frame, bg=self.colors["sidebar_bg"])
        align_row.pack(fill=tk.X, pady=2)
        tk.Label(align_row, text="Align:", bg=self.colors["sidebar_bg"], fg="white").pack(side=tk.LEFT, padx=(2, 0))
        align_combo = ttk.Combobox(align_row, textvariable=self.text_align, values=["left", "center", "right"], width=10)
        align_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        align_combo.bind("<<ComboboxSelected>>", lambda e: self.update_font_lists())

        # BG Image Row
        bg_btn_row = tk.Frame(controls_frame, bg=self.colors["sidebar_bg"])
        bg_btn_row.pack(fill=tk.X, pady=5)
        tk.Button(bg_btn_row, text="Upload BG Image", command=self.upload_bg_image, bg=self.colors["bg"], fg="white", font=("Arial", 8)).pack(fill=tk.X)


    def setup_right_column(self):
        # Notebook (fills everything)
        self.right_notebook = ttk.Notebook(self.right_frame)
        self.right_notebook.pack(fill=tk.BOTH, expand=True)
        self.right_notebook.bind("<<NotebookTabChanged>>", self.on_right_tab_change)
        
        # Show All button (placed beside tabs)
        self.show_all_btn = tk.Button(self.right_frame, text="Show All Fonts", command=self.show_all_fonts, 
                                      bg=self.colors["sidebar_bg"], fg="white", font=("Arial", 8))
        # Initial state hidden
        
        # Font Size Slider (placed beside tabs)
        self.slider_frame = tk.Frame(self.right_frame)
        self.slider_frame.place(relx=1.0, x=-10, y=2, anchor="ne")
        
        tk.Label(self.slider_frame, text="Size:", font=("Arial", 9)).pack(side=tk.LEFT, padx=(0, 2))
        self.size_slider = ttk.Scale(self.slider_frame, from_=8, to_=250, orient=tk.HORIZONTAL,
                                   length=80, command=self.on_size_slider_change)
        self.size_slider.set(self.preview_font_size)
        self.size_slider.pack(side=tk.LEFT, padx=0)
        self.size_slider.bind("<ButtonRelease-1>", self.on_size_slider_release)
        
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


    def on_size_slider_change(self, value):
        new_size = int(float(value))
        if new_size == self.preview_font_size:
            return
        self.preview_font_size = new_size
        
        # Debounce the update to make slider movement smoother
        if self._size_timer:
            self.root.after_cancel(self._size_timer)
        self._size_timer = self.root.after(30, lambda: self.update_font_lists(frequent=True))

    def on_size_slider_release(self, event):
        if self._size_timer:
            self.root.after_cancel(self._size_timer)
            self._size_timer = None
        self.update_font_lists() # Final update to be sure everything is in sync
        self.save_to_history()

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
            "bg_image_path": self.bg_image_path,
            "text_align": self.text_align.get(),
            "preview_text": self.preview_text_box.get("1.0", tk.END),
            "font_bboxes": self.font_bboxes,
            "sort_by_dir": getattr(self, "sort_by_dir", False),
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
        if hasattr(self, 'size_slider'):
            self.size_slider.set(self.preview_font_size)
        self.preview_fg = p["preview_fg"]
        self.preview_bg = p["preview_bg"]
        
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
                self.bg_image_tk = None
        
        self.text_align.set(p["text_align"])
        
        self.preview_text_box.delete("1.0", tk.END)
        self.preview_text_box.insert("1.0", p["preview_text"].strip())
        
        self.font_bboxes = p.get("font_bboxes", {})
        
        self.sort_by_dir = p.get("sort_by_dir", False)
        
        # Synchronize Hex Entries
        self.fg_hex_entry.delete(0, tk.END)
        self.fg_hex_entry.insert(0, self.preview_fg)
        self.bg_hex_entry.delete(0, tk.END)
        self.bg_hex_entry.insert(0, self.preview_bg)
        
        # Update UI components
        self.update_preview()
        self.projects_tree.update_tree()
        
        # Restore tab selections
        if "right_tab_idx" in p and p["right_tab_idx"] < len(self.right_notebook.tabs()):
            self.right_notebook.select(p["right_tab_idx"])

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

    def on_project_right_click(self, event, project_name):
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label=f"Delete Project '{project_name}'", command=lambda: self.delete_project(project_name))
        menu.post(event.x_root, event.y_root)

    def delete_project(self, name):
        if messagebox.askyesno("Delete Project", f"Are you sure you want to delete project '{name}'?"):
            self.save_to_history()
            if name in self.data_manager.data["projects"]:
                del self.data_manager.data["projects"][name]
                if self.current_project == name:
                    self.current_project = None
                self.data_manager.save()
                self.projects_tree.update_tree()

    def add_font_to_project(self, project_name, family):
        p_data = self.data_manager.data["projects"][project_name]
        if family not in p_data:
            self.save_to_history()
            p_data.append(family)
            self.data_manager.save()
            self.projects_tree.update_tree()

    def remove_font_from_project(self, project_name, family):
        p_data = self.data_manager.data["projects"][project_name]
        if family in p_data:
            self.save_to_history()
            p_data.remove(family)
            self.data_manager.save()
            self.projects_tree.update_tree()

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
            self.update_preview()

    def on_bg_hex_change(self, event=None):
        color = self.bg_hex_entry.get()
        if color.startswith("#") and (len(color) == 7 or len(color) == 4):
            if color != self.preview_bg:
                self.save_to_history()
                self.preview_bg = color
                self.update_preview()

    def upload_bg_image(self):
        path = filedialog.askopenfilename(filetypes=[("Image files", "*.jpg *.jpeg *.png *.bmp")])
        if path:
            self.save_to_history()
            self.bg_image_path = path
            try:
                self.bg_image = Image.open(path)
                # Create a PhotoImage for use in Tkinter
                self.bg_image_tk = ImageTk.PhotoImage(self.bg_image)
            except Exception as e:
                print(f"Error loading image: {e}")
                self.bg_image = None
                self.bg_image_tk = None
            self.update_font_lists()

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
        self.sort_by_dir = not getattr(self, "sort_by_dir", False)
        self.update_font_lists()

    def update_font_lists(self, families_filter=None, reset_scroll=False, frequent=False):
        if families_filter is not None:
            self.current_families_filter = families_filter
            self.show_all_btn.place(x=180, y=2)
            reset_scroll = True
        elif not hasattr(self, 'current_families_filter'):
            self.current_families_filter = None

        with self.fonts_lock:
            if self.current_families_filter is not None:
                all_fonts_list = [self.all_fonts[f] for f in self.current_families_filter if f in self.all_fonts]
            else:
                all_fonts_list = list(self.all_fonts.values())
            
        if getattr(self, "sort_by_dir", False):
            sorted_fonts = sorted(all_fonts_list, key=lambda x: (x.directory.lower(), x.family.lower()))
        else:
            sorted_fonts = sorted(all_fonts_list, key=lambda x: x.family.lower())
        
        current_tab = self.right_notebook.index("current") if self.right_notebook.tabs() else 0
        
        # If it's a frequent update (like slider), only update the active tab
        if frequent:
            if current_tab == 0:
                self.all_fonts_list.set_fonts(sorted_fonts)
            else:
                fav_fonts = [f for f in sorted_fonts if f.is_favorite]
                self.fav_fonts_list.set_fonts(fav_fonts)
        else:
            self.all_fonts_list.set_fonts(sorted_fonts)
            fav_fonts = [f for f in sorted_fonts if f.is_favorite]
            self.fav_fonts_list.set_fonts(fav_fonts)
        
        if reset_scroll:
            self.all_fonts_list.reset_scroll()
            self.fav_fonts_list.reset_scroll()

    def show_all_fonts(self):
        self.current_families_filter = None
        self.show_all_btn.place_forget()
        self.update_font_lists(reset_scroll=True)

    def bind_font_drag(self, widget, family):
        widget.bind("<Button-1>", lambda e: self.on_font_drag_start(e, family), add="+")
        widget.bind("<B1-Motion>", self.on_font_drag_motion, add="+")
        widget.bind("<ButtonRelease-1>", self.on_font_drag_release, add="+")

    def on_font_drag_start(self, event, family):
        self.dragged_font = family
        # Create a small floating label for visual feedback
        self.drag_label = tk.Toplevel(self.root)
        self.drag_label.overrideredirect(True)
        self.drag_label.attributes("-topmost", True)
        tk.Label(self.drag_label, text=f"Adding: {family}", bg="#FFD700", fg="black", 
                 padx=5, pady=2, font=("Arial", 9, "bold"), borderwidth=1, relief="solid").pack()
        # Offset to not be directly under the cursor (helps winfo_containing)
        self.drag_label.geometry(f"+{event.x_root + 15}+{event.y_root + 15}")
        self.root.config(cursor="hand2")

    def on_font_drag_motion(self, event):
        if hasattr(self, 'drag_label'):
            self.drag_label.geometry(f"+{event.x_root + 15}+{event.y_root + 15}")

    def on_font_drag_release(self, event):
        self.root.config(cursor="")
        if hasattr(self, 'drag_label'):
            self.drag_label.destroy()
            del self.drag_label

        if not hasattr(self, 'dragged_font'): 
            return
        
        # Get target widget under mouse release position
        x, y = event.x_root, event.y_root
        target = self.root.winfo_containing(x, y)
        
        project_name = None
        curr = target
        while curr:
            if hasattr(curr, 'project_name'):
                project_name = curr.project_name
                break
            if curr == self.root: break
            try:
                curr = curr.master
            except:
                break
            
        if project_name:
            self.add_font_to_project(project_name, self.dragged_font)
        
        del self.dragged_font

    def add_font_to_active_project(self, family):
        if self.current_project:
            self.add_font_to_project(self.current_project, family)
        else:
            messagebox.showinfo("Add Font", "Please select a project first.")

    def on_font_click(self, family):
        self.save_to_history()
        self.last_clicked_font = family
        # We don't call update_preview here because it recreates the entire font list, 
        # which would destroy the widget currently initiating a drag.

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
        self.projects_tree.update_tree()

    def update_preview(self):
        # Update bg_image_tk if bg_image exists and bg_image_tk is None (e.g., after snapshot load)
        if self.bg_image and not self.bg_image_tk:
            try:
                self.bg_image_tk = ImageTk.PhotoImage(self.bg_image)
            except Exception:
                pass
        self.update_font_lists()

    def add_project(self):
        name = simpledialog.askstring("Add Project", "Project Name:")
        if name:
            if name in self.data_manager.data["projects"]:
                messagebox.showerror("Error", "Project already exists.")
                return
            self.save_to_history()
            self.data_manager.data["projects"][name] = []
            self.data_manager.save()
            self.projects_tree.update_tree()

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

    def on_right_tab_change(self, event):
        if not getattr(self, "ignore_history", False):
            if self.right_notebook.tabs():
                self.save_to_history()
        # Always update the newly selected tab to ensure it's in sync
        self.update_font_lists()

if __name__ == "__main__":
    root = tk.Tk()
    app = KorianFontsManagerApp(root)
    root.mainloop()
