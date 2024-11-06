import win32com.client
import os
import math
import time
import json
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog, Text, Toplevel
from PIL import Image, ImageTk, ImageSequence, ImageDraw
import zipfile
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from tkinterdnd2 import TkinterDnD, DND_FILES  # Import DnD2


class ModelAnnotator:
    def __init__(self):
        # Initialize various components
        self.swApp = win32com.client.Dispatch("SldWorks.Application")
        self.model = self.swApp.ActiveDoc
        
        if not self.model:
            messagebox.showerror("Error", "No active SolidWorks document found.")
            exit()
        
        self.frame_notes = {}
        self.frame_markers = {}
        self.current_frame = 0
        self.current_vertical_frame = 0  # Initialize to avoid errors
        self.image_list = []
        self.current_marker = None
        self.marker_count = 0
        self.h_frames = 0  # Initialize frame counters
        self.v_frames = 0
        self.current_photo = None

        self.setup_gui()  # Call setup_gui after initializing attributes

    def setup_gui(self):
        # GUI setup code, such as creating Tkinter windows, goes here.
        # Ensure all elements are created and configured in this function.
        self.root = TkinterDnD.Tk()
        self.root.title("3D Model Annotator")
        self.root.geometry("1200x600")
        # Define frames, canvases, sliders, etc.

    # Add additional methods below setup_gui, like update_progress
    def update_progress(self, current_frame, total_frames):
        # Progress bar update method
        progress = int((current_frame / total_frames) * 100)
        self.progress_bar['value'] = progress
        self.progress_text.config(text=f"{progress}%")
        self.root.update_idletasks()
    

    def on_slider_change(self, value):
        self.current_frame = int(value)
        self.show_frame(self.current_frame, self.current_vertical_frame)  # Adjusted to include vertical frame
        self.frame_counter.config(text=f"Frame: {self.current_frame}/{self.h_frames - 1}")

    def on_vertical_slider_change(self, value):
        self.current_vertical_frame = int(value)
        self.show_frame(self.current_frame, self.current_vertical_frame)

    def setup_gui(self):
        self.root = TkinterDnD.Tk()  # Integrate with DnD2
        self.root.title("3D Model Annotator")
        self.root.geometry("1200x600")  # Set window height to 600px

        # Layout Frames
        self.left_frame = tk.Frame(self.root, width=800, bg="#F0F0F0")
        self.left_frame.grid(row=0, column=0, sticky="nswe")

        self.right_frame = tk.Frame(self.root, width=400)
        self.right_frame.grid(row=0, column=1, sticky="ns")

        self.root.grid_columnconfigure(0, weight=3)
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(0, weight=1)

        # Canvas for Image and Markers
        self.canvas = tk.Canvas(self.left_frame, bg="#F0F0F0", highlightthickness=0)
        self.canvas.pack(expand=True, fill="both", padx=10, pady=10)
        self.canvas.bind("<Button-1>", self.add_marker)

        # Slider Frame with Note Indicators
        self.slider_frame = tk.Frame(self.left_frame)
        self.slider_frame.pack(side="bottom", fill="x", pady=10)

        self.slider = tk.Scale(
            self.slider_frame,
            from_=0,
            to=100,
            orient="horizontal",
            command=self.on_slider_change
        )
        self.slider.pack(fill="x", padx=10, pady=5)

        # Note Indicators Canvas
        self.indicator_canvas = tk.Canvas(self.slider_frame, height=20)
        self.indicator_canvas.pack(fill="x", padx=10)

        # Frame Counter
        self.frame_counter = tk.Label(
            self.slider_frame,
            text="Frame: 0/0",
            font=("Arial", 10)
        )
        self.frame_counter.pack(pady=5)

        # Vertical Slider and Notes Section
        self.vertical_slider_frame = tk.Frame(self.right_frame)
        self.vertical_slider_frame.pack(side="left", fill="y")

        self.vertical_slider = tk.Scale(
            self.vertical_slider_frame,
            from_=0,
            to=100,
            orient="vertical",
            command=self.on_vertical_slider_change
        )
        self.vertical_slider.pack(fill="y", padx=5, pady=5)

        # Notes Section
        self.notes_frame = tk.Frame(self.right_frame)
        self.notes_frame.pack(side="left", fill="both", expand=True)

        notes_label = tk.Label(self.notes_frame, text="NOTES:", anchor="w", pady=5, font=("Arial", 12, "bold"))
        notes_label.pack(fill="x")

        self.notes_listbox = tk.Listbox(self.notes_frame, width=45, height=30)
        self.notes_listbox.pack(expand=True, fill="both", padx=5, pady=5)
        self.notes_listbox.bind('<Double-Button-1>', self.edit_note)
        self.notes_listbox.bind('<Delete>', self.delete_note)
        self.notes_listbox.bind('<<ListboxSelect>>', self.on_note_select)

        # Progress Bar
        self.setup_progress_bar()

        # Menu Configuration
        self.setup_menu()

    def setup_progress_bar(self):
        self.progress_frame = tk.Frame(self.left_frame, bg="#DDDDDD", relief="solid", bd=1)
        self.progress_bar = ttk.Progressbar(
            self.progress_frame,
            orient="horizontal",
            mode="determinate"
        )
        self.progress_bar.pack(expand=True, fill="both", padx=10, pady=5)

        self.progress_text = tk.Label(
            self.progress_frame,
            text="0%",
            bg="#DDDDDD",
            font=("Arial", 14)
        )
        self.progress_text.pack()
        self.progress_frame.place_forget()

    def setup_menu(self):
        menu_bar = tk.Menu(self.root)
        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Process Model", command=self.process_model)
        file_menu.add_command(label="Save as LYNX", command=self.save_lynx)
        file_menu.add_command(label="Open LYNX", command=self.open_lynx)
        file_menu.add_separator()
        file_menu.add_command(label="Save As PDF", command=self.save_as_pdf)
        file_menu.add_command(label="Close", command=self.root.quit)
        menu_bar.add_cascade(label="File", menu=file_menu)
        self.root.config(menu=menu_bar)

    def update_rotation_mode(self):
        pass

    def add_marker(self, event):
        if not self.image_list:
            return

        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)

        self.marker_count += 1

        self.canvas.create_polygon(
            x-10, y+20,
            x, y,
            x+10, y+20,
            fill="red",
            tags=f"marker_{self.marker_count}"
        )

        circle_radius = 10
        self.canvas.create_oval(
            x-circle_radius, y+20-circle_radius,
            x+circle_radius, y+20+circle_radius,
            fill="white",
            tags=f"marker_{self.marker_count}"
        )

        self.canvas.create_text(
            x, y+20,
            text=str(self.marker_count),
            font=("Arial", 8, "bold"),
            tags=f"marker_{self.marker_count}"
        )

        frame_index = self.current_vertical_frame * self.h_frames + self.current_frame
        if frame_index not in self.frame_markers:
            self.frame_markers[frame_index] = []

        self.frame_markers[frame_index].append((x, y, self.marker_count))

        note_text = f"#{self.marker_count}"
        self.notes_listbox.insert(tk.END, note_text)
        self.update_note_indicators()

    def on_note_select(self, event):
        selection = self.notes_listbox.curselection()
        if selection:
            note_text = self.notes_listbox.get(selection[0])
            marker_num = int(note_text.split('-')[0].strip('#').strip())
            frame_num = next(
                (frame for frame, markers in self.frame_markers.items()
                 if any(marker[2] == marker_num for marker in markers)),
                None
            )
            if frame_num is not None:
                v_index = frame_num // self.h_frames
                h_index = frame_num % self.h_frames
                self.slider.set(h_index)
                self.vertical_slider.set(v_index)
                self.show_frame(h_index, v_index)

    def edit_note(self, event):
        selection = self.notes_listbox.curselection()
        if not selection:
            return
        index = selection[0]
        note_text = self.notes_listbox.get(index)

        edit_window = Toplevel(self.root)
        edit_window.title("Edit Note")
        edit_window.geometry("350x300")

        tk.Label(edit_window, text="Enter note details:").pack(anchor="w", padx=10)
        details_text = Text(edit_window, width=40, height=10)
        details_text.pack(padx=10, pady=5)

        author_frame = tk.Frame(edit_window)
        tk.Label(author_frame, text="Author:").pack(side="left", padx=(10, 5))
        author_entry = tk.Entry(author_frame, width=30)
        author_entry.pack(side="left")
        author_frame.pack(anchor="w", pady=(5, 0))

        def save_note():
            details = details_text.get("1.0", tk.END).strip()
            author = author_entry.get().strip()
            updated_text = f"{note_text.split('-')[0]} - {details}"
            if author:
                updated_text += f" (Author: {author})"
            self.notes_listbox.delete(index)
            self.notes_listbox.insert(index, updated_text)
            edit_window.destroy()

        tk.Button(edit_window, text="Save", command=save_note).pack(pady=10)

    def delete_note(self, event):
        selection = self.notes_listbox.curselection()
        if selection:
            note_text = self.notes_listbox.get(selection[0])
            marker_num = int(note_text.split('-')[0].strip('#').strip())

            frame_num = next((frame for frame, markers in self.frame_markers.items()
                              if any(marker[2] == marker_num for marker in markers)), None)
            if frame_num is not None:
                self.frame_markers[frame_num] = [
                    marker for marker in self.frame_markers[frame_num] if marker[2] != marker_num
                ]
                self.canvas.delete(f"marker_{marker_num}")

                if not self.frame_markers[frame_num]:
                    del self.frame_markers[frame_num]
                    self.update_note_indicators()

            self.notes_listbox.delete(selection[0])

    def update_note_indicators(self):
        self.indicator_canvas.delete("all")
        width = self.slider.winfo_width()

        if not self.image_list:
            return

        for frame_num in self.frame_markers.keys():
            h_index = frame_num % self.h_frames
            x_pos = (h_index / (self.h_frames - 1)) * (width - 20) + 10
            self.indicator_canvas.create_polygon(
                x_pos-5, 0,
                x_pos, 10,
                x_pos+5, 0,
                fill="blue"
            )

    def save_as_pdf(self):
        pdf_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not pdf_path:
            return

        pdf = canvas.Canvas(pdf_path, pagesize=A4)

        for frame_num, markers in self.frame_markers.items():
            frame_path = self.image_list[frame_num]
            img = Image.open(frame_path)

            img_draw = ImageDraw.Draw(img)
            for x, y, marker_num in markers:
                img_draw.ellipse((x-5, y-5, x+5, y+5), fill="red")
                img_draw.text((x+10, y), f"#{marker_num}", fill="red")

            img_buffer = io.BytesIO()
            img.save(img_buffer, format="PNG")
            img_buffer.seek(0)

            img_width, img_height = img.size
            aspect_ratio = img_width / img_height
            pdf_width, pdf_height = 300, int(300 / aspect_ratio)

            pdf.drawImage(img_buffer, 50, A4[1] - pdf_height - 100, width=pdf_width, height=pdf_height)
            pdf.drawString(50, A4[1] - pdf_height - 120, f"Frame {frame_num}")

            for marker in markers:
                note_text = self.notes_listbox.get(marker[2] - 1)
                pdf.drawString(50, A4[1] - pdf_height - 140 - (markers.index(marker) * 15), f"#{marker[2]}: {note_text}")

            pdf.showPage()

        pdf.save()
        messagebox.showinfo("Saved", "PDF saved successfully.")

    def process_model(self):
        output_dir = filedialog.askdirectory(title="Select Output Folder")
        if not output_dir:
            return

        self.image_list = self.rotate_and_capture(output_dir)
        if self.image_list:
            self.slider.config(to=self.h_frames - 1)
            self.vertical_slider.config(to=self.v_frames - 1)
            self.show_frame(0, 0)
            self.frame_counter.config(text=f"Frame: 0/{self.h_frames - 1}")

    def save_lynx(self):
        if not self.image_list:
            messagebox.showerror("Error", "No frames to save.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".lynx",
            filetypes=[("LYNX files", "*.lynx")],
            title="Save as LYNX"
        )

        if not file_path:
            return

        try:
            with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as lynx_file:
                webp_data = self.create_webp()
                lynx_file.writestr('animation.webp', webp_data)

                data = {
                    'frame_count': len(self.image_list),
                    'notes': {i: self.notes_listbox.get(i) for i in range(self.notes_listbox.size())},
                    'markers': self.frame_markers
                }
                lynx_file.writestr('data.json', json.dumps(data))

            messagebox.showinfo("Success", "File saved successfully!")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file: {str(e)}")

    def open_lynx(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("LYNX files", "*.lynx")],
            title="Open LYNX File"
        )

        if not file_path:
            return

        try:
            temp_dir = os.path.join(os.path.dirname(file_path), "temp_frames")
            os.makedirs(temp_dir, exist_ok=True)

            with zipfile.ZipFile(file_path, 'r') as lynx_file:
                data = json.loads(lynx_file.read('data.json'))

                self.notes_listbox.delete(0, tk.END)

                for note in data['notes'].values():
                    self.notes_listbox.insert(tk.END, note)

                self.frame_markers = {int(k): v for k, v in data.get('markers', {}).items()}

                webp_data = lynx_file.read('animation.webp')
                webp_buffer = io.BytesIO(webp_data)
                animation = Image.open(webp_buffer)

                self.image_list = []
                for i, frame in enumerate(ImageSequence.Iterator(animation)):
                    frame_path = os.path.join(temp_dir, f"frame_{i:03d}.png")
                    frame.save(frame_path, "PNG")
                    self.image_list.append(frame_path)

                self.slider.config(to=self.h_frames - 1)
                self.vertical_slider.config(to=self.v_frames - 1)
                self.slider.set(0)
                self.vertical_slider.set(0)
                self.show_frame(0, 0)
                self.frame_counter.config(text=f"Frame: 0/{self.h_frames - 1}")
                self.update_note_indicators()

                messagebox.showinfo("Success", "File loaded successfully!")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {str(e)}")
        finally:
            if 'temp_dir' in locals() and os.path.exists(temp_dir):
                try:
                    for file in os.listdir(temp_dir):
                        os.remove(os.path.join(temp_dir, file))
                    os.rmdir(temp_dir)
                except:
                    pass

    def rotate_and_capture(self, output_dir, h_total_degrees=360, h_step_degrees=15, v_total_degrees=180, v_step_degrees=15, delay=0.05):
        self.model.ShowNamedView2("*Top", 1)
        self.model.ViewZoomtofit2()
        time.sleep(delay)

        h_steps = int(h_total_degrees / h_step_degrees)
        v_steps = int(v_total_degrees / v_step_degrees) + 1
        self.h_frames = h_steps
        self.v_frames = v_steps

        self.progress_frame.place(relx=0.5, rely=0.5, anchor="center", relwidth=0.4, relheight=0.1)
        total_frames = h_steps * v_steps
        current_frame = 0

        for v_index, v_angle in enumerate(range(0, v_total_degrees + 1, v_step_degrees)):
            self.model.ActiveView.RotateAboutCenter(math.radians(v_angle), 0)
            time.sleep(delay)

            for h_index in range(h_steps):
                try:
                    self.model.ActiveView.RotateAboutCenter(0, math.radians(h_step_degrees))
                    time.sleep(delay)

                    frame_path = os.path.join(output_dir, f"frame_{current_frame:03d}.png")
                    self.model.SaveAs(frame_path)
                    self.image_list.append(frame_path)

                    current_frame += 1
                    self.update_progress(current_frame, total_frames)
                    time.sleep(delay)
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to capture frame {current_frame}: {e}")
                    self.progress_frame.place_forget()
                    return None

        self.model.ShowNamedView2("*Top", 1)
        self.model.ViewZoomtofit2()
        self.progress_frame.place_forget()
        return self.image_list

    def create_webp(self):
        if not self.image_list:
            return None

        frames = [Image.open(f).convert("RGBA") for f in self.image_list]
        webp_buffer = io.BytesIO()
        frames[0].save(
            webp_buffer,
            format='WEBP',
            save_all=True,
            append_images=frames[1:],
            duration=100,
            loop=0
        )
        return webp_buffer.getvalue()

    def show_frame(self, h_index, v_index):
        frame_index = v_index * self.h_frames + h_index
        if self.image_list:
            if 0 <= frame_index < len(self.image_list):
                self.canvas.delete("all")

                img = Image.open(self.image_list[frame_index])
                canvas_width = self.canvas.winfo_width()
                canvas_height = self.canvas.winfo_height()

                img_ratio = img.size[0] / img.size[1]
                canvas_ratio = canvas_width / canvas_height

                if img_ratio > canvas_ratio:
                    new_width = canvas_width
                    new_height = int(canvas_width / img_ratio)
                else:
                    new_height = canvas_height
                    new_width = int(canvas_height * img_ratio)

                img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                self.current_photo = ImageTk.PhotoImage(img)
                self.canvas.create_image(
                    canvas_width // 2,
                    canvas_height // 2,
                    image=self.current_photo,
                    anchor="center"
                )

                if frame_index in self.frame_markers:
                    for x, y, marker_num in self.frame_markers[frame_index]:
                        self.canvas.create_polygon(
                            x-10, y+20,
                            x, y,
                            x+10, y+20,
                            fill="red",
                            tags=f"marker_{marker_num}"
                        )

                        circle_radius = 10
                        self.canvas.create_oval(
                            x-circle_radius, y+20-circle_radius,
                            x+circle_radius, y+20+circle_radius,
                            fill="white",
                            tags=f"marker_{marker_num}"
                        )

                        self.canvas.create_text(
                            x, y+20,
                            text=str(marker_num),
                            font=("Arial", 8, "bold"),
                            tags=f"marker_{marker_num}"
                        )

if __name__ == "__main__":
    app = ModelAnnotator()
    app.root.mainloop()
