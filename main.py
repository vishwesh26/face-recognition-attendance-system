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
        
        # Add icon
        try:
            self.root.iconbitmap("assets/icon.ico")
        except:
            pass
        
        # Configure the root window
        self.root.configure(background='#f0f0f0')
        
        # Create directories if they don't exist
        directories = ['TrainingImages', 'TrainingImageLabels', 'AttendanceRecords', 'assets']
        for dir in directories:
            if not os.path.exists(dir):
                os.makedirs(dir)
        
        # Variables
        self.var_student_name = tk.StringVar()
        self.var_roll_no = tk.StringVar()
        self.var_subject = tk.StringVar()
        self.var_email = tk.StringVar()
        self.var_email_enabled = tk.BooleanVar(value=True)
        
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
                bg="#f0f0f0").grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
        
        tk.Entry(left_frame, textvariable=self.var_student_name, 
                font=("Helvetica", 12), width=20).grid(row=0, column=1, padx=10, pady=10, sticky=tk.W)
        
        # Roll Number
        tk.Label(left_frame, text="Roll Number:", font=("Helvetica", 12), 
                bg="#f0f0f0").grid(row=1, column=0, padx=10, pady=10, sticky=tk.W)
        
        tk.Entry(left_frame, textvariable=self.var_roll_no, 
                font=("Helvetica", 12), width=20).grid(row=1, column=1, padx=10, pady=10, sticky=tk.W)
        
        # Email
        tk.Label(left_frame, text="Email Address:", font=("Helvetica", 12), 
                bg="#f0f0f0").grid(row=2, column=0, padx=10, pady=10, sticky=tk.W)
        
        tk.Entry(left_frame, textvariable=self.var_email, 
                font=("Helvetica", 12), width=20).grid(row=2, column=1, padx=10, pady=10, sticky=tk.W)
        
        # Subject
        tk.Label(left_frame, text="Subject:", font=("Helvetica", 12), 
                bg="#f0f0f0").grid(row=3, column=0, padx=10, pady=10, sticky=tk.W)
        
        subjects = ["Mathematics", "Physics", "Chemistry", "Biology", "Computer Science"]
        subject_combo = ttk.Combobox(left_frame, textvariable=self.var_subject, 
                                   font=("Helvetica", 12), width=18, state="readonly")
        subject_combo["values"] = subjects
        subject_combo.current(0)
        subject_combo.grid(row=3, column=1, padx=10, pady=10, sticky=tk.W)
        
        # Email notifications checkbox
        email_check = tk.Checkbutton(left_frame, text="Send Email Notifications", 
                                   variable=self.var_email_enabled, 
                                   font=("Helvetica", 12), bg="#f0f0f0")
        email_check.grid(row=4, column=0, padx=10, pady=10, columnspan=2, sticky=tk.W)
        
        # Camera Preview
        self.preview_frame = tk.LabelFrame(left_frame, text="Camera Preview", 
                                         font=("Helvetica", 10), bg="#f0f0f0", fg="#4a6ea9")
        self.preview_frame.place(x=10, y=220, width=440, height=220)
        
        self.preview_label = tk.Label(self.preview_frame, bg="black")
        self.preview_label.place(x=0, y=0, width=440, height=195)

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
        instruction_text.insert(tk.END, "1. Register a new student with name, roll number and email\n"
                                      "2. Click 'Capture Images' to take 50 sample images\n"
                                      "3. Train the model after adding multiple students\n"
                                      "4. Select a subject for attendance tracking\n"
                                      "5. Click 'Take Attendance' to begin face recognition\n"
                                      "6. Attendance records will be sent via email if enabled\n"
                                      "7. View attendance reports in Excel format\n")
        instruction_text.configure(state="disabled")
        
        # Buttons
        btn_frame = tk.Frame(right_frame, bg="#f0f0f0")
        btn_frame.place(x=10, y=250, width=450, height=200)
        
        buttons = [
            ("Capture Images", self.capture_images),
            ("Train Model", self.train_model),
            ("Take Attendance", self.take_attendance),
            ("View Attendance", self.view_attendance),
            ("Send Reports", self.send_reports),
            ("Exit", self.exit_program)
        ]
        
        for i, (text, command) in enumerate(buttons):
            btn = tk.Button(btn_frame, text=text, command=command,
                          font=("Helvetica", 12, "bold"), bg="#4a6ea9", fg="white",
                          activebackground="#36518a", activeforeground="white",
                          cursor="hand2", width=15, height=2)
            btn.grid(row=i//3, column=i%3, padx=2, pady=5)

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
            student_emails = atm.get_student_emails()
            
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
                            marked = self._mark_attendance(id, name)
                            
                            # Send email if enabled and attendance marked
                            if marked and self.var_email_enabled.get() and id in student_emails:
                                current_date = datetime.now().strftime('%Y-%m-%d')
                                atm.send_attendance_email(
                                    student_emails[id], 
                                    name,
                                    self.var_subject.get(),
                                    current_date
                                )
                            
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
            file_path = os.path.abspath(f"AttendanceRecords/{self.var_subject.get()}_attendance.xlsx")
            
            if not os.path.exists(file_path):
                raise Exception(f"No attendance records for {self.var_subject.get()}")
            
            os.startfile(file_path) if os.name == 'nt' else os.system(f"xdg-open {file_path}")
            self.status_var.set(f"Status: Viewing attendance for {self.var_subject.get()}")
            
        except Exception as e:
            self._handle_error("View Attendance", str(e))
    
    def send_reports(self):
        """Send attendance reports to all students"""
        if not self._validate_subject():
            return
        
        try:
            file_path = f"AttendanceRecords/{self.var_subject.get()}_attendance.xlsx"
            
            if not os.path.exists(file_path):
                raise Exception(f"No attendance records for {self.var_subject.get()}")
            
            # Load student emails
            student_emails = atm.get_student_emails()
            if not student_emails:
                raise Exception("No student email records found")
            
            # Send emails to all students
            current_date = datetime.now().strftime('%Y-%m-%d')
            email_sent_count = 0
            
            for student_id, email in student_emails.items():
                student_data = self._load_student_data()
                name = student_data.get(student_id, f"Student {student_id}")
                
                if atm.send_attendance_email(email, name, self.var_subject.get(), current_date):
                    email_sent_count += 1
            
            self.status_var.set(f"Status: Reports sent to {email_sent_count} students")
            messagebox.showinfo("Email Sent", f"Attendance reports sent to {email_sent_count} students")
            
        except Exception as e:
            self._handle_error("Send Reports", str(e))
    
    def exit_program(self):
        """Exit the program"""
        self._cleanup_camera()
        if messagebox.askyesno("Exit", "Are you sure you want to exit?"):
            self.root.destroy()

    def _validate_inputs(self):
        """Validate student inputs"""
        if not self.var_student_name.get() or not self.var_roll_no.get():
            messagebox.showerror("Error", "Please enter name and roll number")
            return False
        
        if self.var_email_enabled.get() and not self.var_email.get():
            messagebox.showerror("Error", "Please enter email address or disable email notifications")
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
        # Save to StudentDetails.csv
        with open("TrainingImageLabels/StudentDetails.csv", 'a+') as f:
            f.seek(0)
            existing_rolls = [line.split(',')[0] for line in f.readlines() if line.strip()]
            
            if self.var_roll_no.get() in existing_rolls:
                raise Exception("Roll number already exists!")
            
            f.write(f"{self.var_roll_no.get()},{self.var_student_name.get()}\n")
        
        # Save email if provided
        if self.var_email.get():
            email_file = "TrainingImageLabels/StudentEmails.csv"
            
            # Create or append to email file
            if not os.path.exists(email_file):
                with open(email_file, 'w') as f:
                    f.write("ID,Email\n")
            
            with open(email_file, 'a') as f:
                f.write(f"{self.var_roll_no.get()},{self.var_email.get()}\n")

    def _load_student_data(self):
        """Load student data from CSV"""
        student_data = {}
        try:
            with open("TrainingImageLabels/StudentDetails.csv", 'r') as f:
                for line in f:
                    if line.strip():
                        parts = line.strip().split(',')
                        if len(parts) >= 2:
                            roll, name = parts[0], parts[1]
                            student_data[int(roll)] = name
        except Exception as e:
            print(f"Error loading student data: {str(e)}")
        
        return student_data

    def _mark_attendance(self, student_id, name):
        """Mark attendance for a student"""
        date = datetime.now().strftime('%Y-%m-%d')
        time = datetime.now().strftime('%H:%M:%S')
        
        marked = atm.mark_attendance(student_id, name, self.var_subject.get(), date, time)
        if marked:
            self.status_var.set(f"Status: Attendance marked for {name}")
        
        return marked

    def _update_preview(self, frame):
        """Update the preview label with the current frame"""
        frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        img = Image.fromarray(frame)
        img = ImageTk.PhotoImage(image=img)
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