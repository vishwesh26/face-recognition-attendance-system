import tkinter as tk
from tkinter import ttk, messagebox
import cv2
import os
import numpy as np
from datetime import datetime
from PIL import Image, ImageTk
import face_recognition_module as frm
import attendance_module as atm

class FaceAttendanceSystem:
    def __init__(self, root):
        self.root = root
        self.root.title("Face Recognition Attendance System")
        
        # Set window size and center it
        window_width = 1000
        window_height = 600
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        center_x = int(screen_width/2 - window_width/2)
        center_y = int(screen_height/2 - window_height/2)
        self.root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        self.root.resizable(False, False)
        
        # Configure the root window
        self.root.configure(background='#f0f0f0')
        
        # Create directories if they don't exist
        directories = ['TrainingImages', 'TrainingImageLabels', 'AttendanceRecords']
        for dir in directories:
            if not os.path.exists(dir):
                os.makedirs(dir)
        
        # Variables
        self.var_student_name = tk.StringVar()
        self.var_roll_no = tk.StringVar()
        self.var_subject = tk.StringVar()
        
        # Create GUI elements
        self._create_gui()
        
        # Initialize webcam
        self.cap = None
        
    def _create_gui(self):
        """Create all GUI elements"""
        # Title Frame
        title_frame = tk.Frame(self.root, bg="#4a6ea9")
        title_frame.place(x=0, y=0, width=1000, height=60)
        
        title_label = tk.Label(title_frame, text="Face Recognition Attendance System", 
                             font=("Helvetica", 20, "bold"), bg="#4a6ea9", fg="white")
        title_label.pack(side=tk.LEFT, padx=20)
        
        # Main Frame
        main_frame = tk.Frame(self.root, bg="#f0f0f0")
        main_frame.place(x=0, y=60, width=1000, height=540)
        
        # Left Frame (Registration)
        self._create_left_frame(main_frame)
        
        # Right Frame (Operations)
        self._create_right_frame(main_frame)
        
        # Status Bar
        self.status_var = tk.StringVar()
        self.status_var.set("Status: Ready")
        status_bar = tk.Label(self.root, textvariable=self.status_var, 
                            font=("Helvetica", 10), bd=1, relief=tk.SUNKEN, anchor=tk.W,
                            bg="#4a6ea9", fg="white")
        status_bar.place(x=0, y=580, width=1000, height=20)

    def _create_left_frame(self, parent):
        """Create left frame with registration controls"""
        left_frame = tk.LabelFrame(parent, text="Student Registration", 
                                 font=("Helvetica", 12, "bold"), bg="#f0f0f0", fg="#4a6ea9")
        left_frame.place(x=20, y=20, width=470, height=500)
        
        # Student Name
        tk.Label(left_frame, text="Student Name:", font=("Helvetica", 12), 
                bg="#f0f0f0").grid(row=0, column=0, padx=10, pady=20, sticky=tk.W)
        
        tk.Entry(left_frame, textvariable=self.var_student_name, 
                font=("Helvetica", 12), width=20).grid(row=0, column=1, padx=10, pady=20, sticky=tk.W)
        
        # Roll Number
        tk.Label(left_frame, text="Roll Number:", font=("Helvetica", 12), 
                bg="#f0f0f0").grid(row=1, column=0, padx=10, pady=20, sticky=tk.W)
        
        tk.Entry(left_frame, textvariable=self.var_roll_no, 
                font=("Helvetica", 12), width=20).grid(row=1, column=1, padx=10, pady=20, sticky=tk.W)
        
        # Subject
        tk.Label(left_frame, text="Subject:", font=("Helvetica", 12), 
                bg="#f0f0f0").grid(row=2, column=0, padx=10, pady=20, sticky=tk.W)
        
        subjects = ["Mathematics", "Physics", "Chemistry", "Biology", "Computer Science"]
        subject_combo = ttk.Combobox(left_frame, textvariable=self.var_subject, 
                                   font=("Helvetica", 12), width=18, state="readonly")
        subject_combo["values"] = subjects
        subject_combo.current(0)
        subject_combo.grid(row=2, column=1, padx=10, pady=20, sticky=tk.W)
        
        # Camera Preview
        self.preview_frame = tk.LabelFrame(left_frame, text="Camera Preview", 
                                         font=("Helvetica", 10), bg="#f0f0f0", fg="#4a6ea9")
        self.preview_frame.place(x=10, y=200, width=440, height=300)
        
        self.preview_label = tk.Label(self.preview_frame, bg="black")
        self.preview_label.place(x=0, y=0, width=440, height=300)

    def _create_right_frame(self, parent):
        """Create right frame with operation buttons"""
        right_frame = tk.LabelFrame(parent, text="Operations", 
                                  font=("Helvetica", 12, "bold"), bg="#f0f0f0", fg="#4a6ea9")
        right_frame.place(x=510, y=20, width=470, height=500)
        
        # Instructions
        tk.Label(right_frame, text="System Instructions", 
                font=("Helvetica", 12, "bold"), bg="#f0f0f0", fg="#4a6ea9").place(x=150, y=20)
        
        instruction_text = tk.Text(right_frame, font=("Helvetica", 10), bg="#ffffff", width=52, height=10)
        instruction_text.place(x=10, y=50)
        instruction_text.insert(tk.END, "1. Register a new student with name and roll number\n"
                                      "2. Click 'Capture Images' to take 50 sample images\n"
                                      "3. Train the model after adding multiple students\n"
                                      "4. Select a subject for attendance tracking\n"
                                      "5. Click 'Take Attendance' to begin face recognition\n"
                                      "6. View attendance reports in Excel format\n")
        instruction_text.configure(state="disabled")
        
        # Buttons
        btn_frame = tk.Frame(right_frame, bg="#f0f0f0")
        btn_frame.place(x=10, y=250, width=450, height=200)
        
        buttons = [
            ("Capture Images", self.capture_images),
            ("Train Model", self.train_model),
            ("Take Attendance", self.take_attendance),
            ("View Attendance", self.view_attendance)
        ]
        
        for i, (text, command) in enumerate(buttons):
            btn = tk.Button(btn_frame, text=text, command=command,
                          font=("Helvetica", 12, "bold"), bg="#4a6ea9", fg="white",
                          activebackground="#36518a", activeforeground="white",
                          cursor="hand2", width=20, height=2)
            btn.grid(row=i//2, column=i%2, padx=25, pady=15)

    def capture_images(self):
        """Capture training images for a student"""
        if not self._validate_inputs():
            return
        
        try:
            self.cap = cv2.VideoCapture(0)
            self.status_var.set("Status: Capturing images...")
            
            path = f"TrainingImages/{self.var_roll_no.get()}"
            os.makedirs(path, exist_ok=True)
            
            # Save student info
            self._save_student_info()
            
            count = 0
            face_detector = cv2.CascadeClassifier(
                cv2.data.haarcascades + 'haarcascade_frontalface_default.xml'
            )
            
            while count < 50:
                ret, frame = self.cap.read()
                if not ret:
                    raise Exception("Cannot access webcam")
                
                gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                faces = face_detector.detectMultiScale(gray, 1.3, 5)
                
                for (x, y, w, h) in faces:
                    cv2.rectangle(frame, (x, y), (x+w, y+h), (255, 0, 0), 2)
                    cv2.imwrite(f"{path}/{count}.jpg", gray[y:y+h, x:x+w])
                    count += 1
                
                # Display progress
                cv2.putText(frame, f"Captured: {count}/50", (50, 50), 
                          cv2.FONT_HERSHEY_COMPLEX, 0.8, (0, 255, 0), 2)
                
                # Update preview
                self._update_preview(frame)
                
                if cv2.waitKey(1) == 27 or count >= 50:  # ESC key
                    break
            
            self._cleanup_camera()
            messagebox.showinfo("Success", 
                              f"Successfully captured {count} images for {self.var_student_name.get()}")
            
        except Exception as e:
            self._handle_error("Image Capture", str(e))

    def train_model(self):
        """Train the face recognition model"""
        try:
            self.status_var.set("Status: Training model...")
            
            if not os.path.exists("TrainingImages") or len(os.listdir("TrainingImages")) == 0:
                raise Exception("No images found for training")
            
            faces, ids = frm.get_faces_and_ids("TrainingImages")
            
            if len(faces) == 0:
                raise Exception("No faces detected in training images")
            
            recognizer = cv2.face.LBPHFaceRecognizer_create()
            recognizer.train(faces, np.array(ids))
            recognizer.write("trainer.yml")
            
            self.status_var.set(f"Status: Model trained successfully with {len(faces)} faces")
            messagebox.showinfo("Success", f"Model trained successfully with {len(faces)} faces")
            
        except Exception as e:
            self._handle_error("Model Training", str(e))

    def take_attendance(self):
        """Take attendance using face recognition"""
        if not self._validate_subject():
            return
            
        try:
            if not os.path.exists("trainer.yml"):
                raise Exception("Model not trained yet!")
            
            student_data = self._load_student_data()
            
            self.cap = cv2.VideoCapture(0)
            self.status_var.set("Status: Taking attendance...")
            
            recognizer = cv2.face.LBPHFaceRecognizer_create()
            recognizer.read("trainer.yml")
            
            face_cascade = cv2.CascadeClassifier(
                cv2.data.haarcascades + 'haarcascade_frontalface_default.xml'
            )
            
            attendance_tracker = set()
            
            while True:
                ret, frame = self.cap.read()
                if not ret:
                    raise Exception("Cannot access webcam")
                
                gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                faces = face_cascade.detectMultiScale(gray, 1.3, 5)
                
                for (x, y, w, h) in faces:
                    cv2.rectangle(frame, (x, y), (x+w, y+h), (255, 0, 0), 2)
                    
                    id, confidence = recognizer.predict(gray[y:y+h, x:x+w])
                    
                    if confidence < 80:
                        name = student_data.get(id, "Unknown")
                        if id not in attendance_tracker:
                            self._mark_attendance(id, name)
                            attendance_tracker.add(id)
                    else:
                        name = "Unknown"
                    
                    # Display name and confidence
                    cv2.putText(frame, f"{name} ({100-confidence:.1f}%)", 
                              (x, y-10), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (255, 255, 255), 2)
                
                self._update_preview(frame)
                
                if cv2.waitKey(1) == ord('q'):
                    break
            
            self._cleanup_camera()
            
        except Exception as e:
            self._handle_error("Attendance", str(e))

    def view_attendance(self):
        """View attendance records"""
        if not self._validate_subject():
            return
            
        try:
            file_path = f"AttendanceRecords/{self.var_subject.get()}_attendance.xlsx"
            
            if not os.path.exists(file_path):
                raise Exception(f"No attendance records for {self.var_subject.get()}")
            
            os.startfile(file_path) if os.name == 'nt' else os.system(f"xdg-open {file_path}")
            self.status_var.set(f"Status: Viewing attendance for {self.var_subject.get()}")
            
        except Exception as e:
            self._handle_error("View Attendance", str(e))

    def _validate_inputs(self):
        """Validate student inputs"""
        if not self.var_student_name.get() or not self.var_roll_no.get():
            messagebox.showerror("Error", "Please enter name and roll number")
            return False
        return True

    def _validate_subject(self):
        """Validate subject selection"""
        if not self.var_subject.get():
            messagebox.showerror("Error", "Please select a subject")
            return False
        return True

    def _save_student_info(self):
        """Save student information to CSV"""
        with open("TrainingImageLabels/StudentDetails.csv", 'a+') as f:
            f.seek(0)
            existing_rolls = [line.split(',')[0] for line in f.readlines()]
            
            if self.var_roll_no.get() in existing_rolls:
                raise Exception("Roll number already exists!")
            
            f.write(f"{self.var_roll_no.get()},{self.var_student_name.get()}\n")

    def _load_student_data(self):
        """Load student data from CSV"""
        student_data = {}
        with open("TrainingImageLabels/StudentDetails.csv", 'r') as f:
            for line in f:
                roll, name = line.strip().split(',')
                student_data[int(roll)] = name
        return student_data

    def _mark_attendance(self, student_id, name):
        """Mark attendance for a student"""
        date = datetime.now().strftime('%Y-%m-%d')
        time = datetime.now().strftime('%H:%M:%S')
        
        if atm.mark_attendance(student_id, name, self.var_subject.get(), date, time):
            self.status_var.set(f"Status: Attendance marked for {name}")

    def _update_preview(self, frame):
        """Update the preview label with the current frame"""
        frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        img = Image.fromarray(frame)
        img = ImageTk.PhotoImage(img)
        self.preview_label.configure(image=img)
        self.preview_label.image = img
        self.root.update()

    def _cleanup_camera(self):
        """Release camera resources"""
        if self.cap:
            self.cap.release()
            self.preview_label.configure(image=None)
            self.preview_label.image = None

    def _handle_error(self, operation, error_msg):
        """Handle errors and update status"""
        self.status_var.set(f"Status: Error during {operation}")
        messagebox.showerror("Error", f"An error occurred during {operation}: {error_msg}")
        self._cleanup_camera()

if __name__ == "__main__":
    root = tk.Tk()
    app = FaceAttendanceSystem(root)
    root.mainloop()