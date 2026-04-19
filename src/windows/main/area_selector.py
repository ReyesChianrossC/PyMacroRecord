from tkinter import *
import tkinter as tk

class AreaSelector(Toplevel):
    def __init__(self, parent, callback):
        super().__init__(parent)
        self.callback = callback
        # Use overrideredirect for borderless fullscreen
        self.overrideredirect(True)
        self.geometry(f"{self.winfo_screenwidth()}x{self.winfo_screenheight()}+0+0")
        self.attributes('-topmost', True)
        self.attributes('-alpha', 0.3)
        self.configure(bg='black')
        
        self.start_x = None
        self.start_y = None
        
        self.canvas = Canvas(self, cursor="cross", bg="grey", highlightthickness=0)
        self.canvas.pack(fill=BOTH, expand=True)
        
        # Bind events for Drag-and-Drop
        self.canvas.bind("<ButtonPress-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)
        self.bind("<Escape>", self.cancel_selection)
        
        self.rect_id = None
        
        # Add instruction label
        self.instruction_text = "Click and Drag to select area (Esc to cancel)"
        self.label = Label(self, text=self.instruction_text, font=("Arial", 20), fg="white", bg="black")
        self.label.place(relx=0.5, rely=0.1, anchor=CENTER)
        
        # Add Cancel Button as failsafe
        self.cancel_btn = Button(self, text="Cancel (Esc)", command=self.cancel_selection, bg="red", fg="white", font=("Arial", 12))
        self.cancel_btn.place(relx=0.9, rely=0.05, anchor=NE)

        # Force focus
        self.update_idletasks()
        self.focus_set()
        # Don't use grab_set - it blocks the messagebox dialog

    def cancel_selection(self, event=None):
        self.callback(None)
        self.destroy()

    def on_press(self, event):
        self.start_x = event.x
        self.start_y = event.y
        if self.rect_id:
            self.canvas.delete(self.rect_id)
        self.rect_id = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline="red", width=3)

    def on_drag(self, event):
        if self.start_x is not None:
            self.canvas.coords(self.rect_id, self.start_x, self.start_y, event.x, event.y)

    def on_release(self, event):
        if self.start_x is not None:
            x1, y1 = self.start_x, self.start_y
            x2, y2 = event.x, event.y
            
            # Normalize coordinates (ensure x1 < x2, y1 < y2)
            left = min(x1, x2)
            top = min(y1, y2)
            right = max(x1, x2)
            bottom = max(y1, y2)
            
            # Helper: prevent accidental 0-size clicks
            if right - left < 10 or bottom - top < 10:
                # Reset if too small (likely a click, not a drag)
                self.start_x = None
                self.start_y = None
                if self.rect_id:
                    self.canvas.delete(self.rect_id)
                self.rect_id = None
                return

            area = [left, top, right, bottom]
            self.callback(area)
            self.destroy()


class MultiAreaSelector(Toplevel):
    """
    Multi-area selector: stays open after each drag so the user can define
    multiple scan regions. Each completed drag leaves a numbered rectangle
    on screen. 'Done' finalises and calls callback([[x1,y1,x2,y2], ...]).
    'Cancel / Esc' calls callback([]) and closes.
    """

    def __init__(self, parent, callback):
        super().__init__(parent)
        self.callback = callback

        self.overrideredirect(True)
        self.geometry(f"{self.winfo_screenwidth()}x{self.winfo_screenheight()}+0+0")
        self.attributes('-topmost', True)
        self.attributes('-alpha', 0.35)
        self.configure(bg='black')

        self.start_x = None
        self.start_y = None
        self.rect_id = None      # The live rubber-band rect while dragging
        self.areas = []          # Finalized [[x1,y1,x2,y2], ...]
        self.area_rects = []     # Canvas IDs of finalized rects
        self.area_labels = []    # Canvas IDs of number labels

        self.canvas = Canvas(self, cursor="cross", bg="grey15", highlightthickness=0)
        self.canvas.pack(fill=BOTH, expand=True)

        self.canvas.bind("<ButtonPress-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)
        self.bind("<Escape>", self.cancel_selection)

        # Instruction label (top centre)
        self.instr_label = Label(
            self,
            text="Drag areas to monitor. Click [Done] when finished.  Esc = cancel.",
            font=("Arial", 18, "bold"),
            fg="yellow", bg="black"
        )
        self.instr_label.place(relx=0.5, rely=0.07, anchor=CENTER)

        # Counter label
        self.count_label = Label(
            self,
            text="Areas: 0",
            font=("Arial", 14),
            fg="white", bg="black"
        )
        self.count_label.place(relx=0.5, rely=0.13, anchor=CENTER)

        # Done button (top right)
        self.done_btn = Button(
            self,
            text="✔ Done",
            command=self.finish_selection,
            bg="#2ecc71", fg="white",
            font=("Arial", 13, "bold"),
            relief=FLAT, padx=10, pady=4
        )
        self.done_btn.place(relx=0.9, rely=0.05, anchor=NE)

        # Cancel button (top right, below Done)
        self.cancel_btn = Button(
            self,
            text="✖ Cancel (Esc)",
            command=self.cancel_selection,
            bg="#e74c3c", fg="white",
            font=("Arial", 12),
            relief=FLAT, padx=6, pady=4
        )
        self.cancel_btn.place(relx=0.9, rely=0.11, anchor=NE)

        # Undo last area button
        self.undo_btn = Button(
            self,
            text="↩ Undo Last",
            command=self.undo_last,
            bg="#e67e22", fg="white",
            font=("Arial", 12),
            relief=FLAT, padx=6, pady=4
        )
        self.undo_btn.place(relx=0.9, rely=0.17, anchor=NE)

        self.update_idletasks()
        self.focus_set()

    # ── drag logic ────────────────────────────────────────────────────────────

    def on_press(self, event):
        self.start_x = event.x
        self.start_y = event.y
        if self.rect_id:
            self.canvas.delete(self.rect_id)
        self.rect_id = self.canvas.create_rectangle(
            self.start_x, self.start_y, self.start_x, self.start_y,
            outline="cyan", width=2, dash=(4, 2)
        )

    def on_drag(self, event):
        if self.start_x is not None:
            self.canvas.coords(self.rect_id, self.start_x, self.start_y, event.x, event.y)

    def on_release(self, event):
        if self.start_x is None:
            return

        x1, y1 = self.start_x, self.start_y
        x2, y2 = event.x, event.y

        left   = min(x1, x2)
        top    = min(y1, y2)
        right  = max(x1, x2)
        bottom = max(y1, y2)

        # Remove the rubber-band rect
        if self.rect_id:
            self.canvas.delete(self.rect_id)
            self.rect_id = None
        self.start_x = None
        self.start_y = None

        # Ignore tiny accidental clicks
        if right - left < 10 or bottom - top < 10:
            return

        self.areas.append([left, top, right, bottom])
        n = len(self.areas)

        # Draw a permanent coloured rectangle
        colors = ["#e74c3c", "#3498db", "#2ecc71", "#f39c12",
                  "#9b59b6", "#1abc9c", "#e67e22", "#34495e"]
        colour = colors[(n - 1) % len(colors)]
        r_id = self.canvas.create_rectangle(
            left, top, right, bottom,
            outline=colour, width=3
        )
        lbl_id = self.canvas.create_text(
            left + 6, top + 6,
            text=str(n),
            fill=colour,
            font=("Arial", 14, "bold"),
            anchor=NW
        )
        self.area_rects.append(r_id)
        self.area_labels.append(lbl_id)
        self.count_label.config(text=f"Areas: {n}")

    # ── buttons ───────────────────────────────────────────────────────────────

    def undo_last(self):
        if not self.areas:
            return
        self.areas.pop()
        if self.area_rects:
            self.canvas.delete(self.area_rects.pop())
        if self.area_labels:
            self.canvas.delete(self.area_labels.pop())
        self.count_label.config(text=f"Areas: {len(self.areas)}")

    def finish_selection(self):
        """Confirm — passes the list of areas (may be empty) back."""
        result = list(self.areas)
        self.destroy()
        self.callback(result)

    def cancel_selection(self, event=None):
        self.destroy()
        self.callback([])
