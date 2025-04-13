import cv2
import os
import numpy as np

def get_faces_and_ids(path):
    """
    Extract faces and their corresponding IDs from the training images
    """
    face_samples = []
    ids = []
    detector = cv2.CascadeClassifier(cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')
    
    # Get all subdirectories with student images
    student_dirs = [os.path.join(path, d) for d in os.listdir(path) if os.path.isdir(os.path.join(path, d))]
    
    for student_dir in student_dirs:
        student_id = int(os.path.basename(student_dir))
        
        # Process all images in student directory
        for img_file in os.listdir(student_dir):
            img_path = os.path.join(student_dir, img_file)
            
            # Read image and convert to grayscale
            img = cv2.imread(img_path)
            if img is None:
                continue
                
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            
            # Detect faces
            faces = detector.detectMultiScale(gray, 1.3, 5)
            
            for (x, y, w, h) in faces:
                face_samples.append(gray[y:y+h, x:x+w])
                ids.append(student_id)
    
    return face_samples, ids