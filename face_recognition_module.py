import cv2
import os
import numpy as np

def get_faces_and_ids(directory):
    """
    Extract faces and their IDs from images in the training directory
    
    Args:
        directory (str): Path to the directory containing training images
        
    Returns:
        tuple: List of face samples and list of IDs
    """
    face_samples = []
    ids = []
    
    # Load face detector
    detector = cv2.CascadeClassifier(cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')
    
    # Navigate through all directories
    for student_id in os.listdir(directory):
        student_path = os.path.join(directory, student_id)
        
        # Check if this is a directory
        if not os.path.isdir(student_path):
            continue
            
        # Process each image file in the directory
        for img_file in os.listdir(student_path):
            if img_file.endswith(('.png', '.jpg', '.jpeg')):
                img_path = os.path.join(student_path, img_file)
                
                # Read the image
                img = cv2.imread(img_path)
                gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
                
                # Detect faces in the image
                faces = detector.detectMultiScale(gray, 1.3, 5)
                
                # If faces are detected, add them to the training set
                for (x, y, w, h) in faces:
                    face_samples.append(gray[y:y+h, x:x+w])
                    ids.append(int(student_id))
    
    return face_samples, ids