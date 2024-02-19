import cv2
import pytesseract
from tkinter import Tk, filedialog, Label, Button, Canvas, Scrollbar, Text, Frame, Entry
from PIL import Image, ImageTk
import pandas as pd
import re

# Tesseract OCR'ın dizinini belirtin
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def process_image():
    global output_list, seen_names

    # Resmi okuma
    image = cv2.imread(image_path)

    # Resmi gri tonlara çevirme
    gray_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

    # Adaptive thresholding uygulama
    _, threshold_image = cv2.threshold(gray_image, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

    # Kontur tespiti
    contours, _ = cv2.findContours(threshold_image, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    # Dataset.xlsx dosyasını okuma
    df = pd.read_excel('Dataset.xlsx')

    # Çıktıları depolamak için bir liste oluştur
    output_list = []

    # Extracted Word'leri içeren kelimeleri ara ve eğer bulunursa ilgili bilgileri ekle
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        roi = image[y:y + h, x:x + w]
        text = pytesseract.image_to_string(roi, config='--psm 6 -l eng')

        sentences = re.split(r'[.!?]', text)
        for sentence in sentences:
            words = re.findall(r'\b\w+\b', sentence)
            for word in words:
                if len(word) > 2:
                    # Kelimenin "Chemical Substance" sütununda bulunup bulunmadığını kontrol et
                    if any(df[" Chemical Substance "].str.contains(word, case=False)):
                        # Kelimenin bulunduğu satırı bul
                        matching_rows = df[df[" Chemical Substance "].str.contains(word, case=False)]
                        for index, row in matching_rows.iterrows():
                            # İlgili bilgileri çıktı listesine ekle
                            output_list.append({
                                "Name": row[" Chemical Substance "],
                                "Health Point": row[" Is it Harmful? "],
                                "What's the use?": row[" Description of the Substance "]
                            })

    # Çiftleri kontrol ederek çiftleri bir kere yazdır
    seen_names = set()
    output_text = ""
    for output in output_list:
        if output["Name"] not in seen_names:
            output_text += "Name: {}\nHealth Point: {}\nWhat's the use?: {}\n\n".format(
                output['Name'], output['Health Point'], output['What\'s the use?'])
            seen_names.add(output["Name"])

    # Görselleştirmeleri ekle
    result_text.config(state="normal")
    result_text.delete(1.0, "end")
    result_text.insert("insert", output_text)
    result_text.config(state="disabled")
    display_selected_image(image_path)

def open_file_dialog():
    global image_path
    image_path = filedialog.askopenfilename()
    path_label.config(text=f"Image Path: {image_path}")
    display_selected_image(image_path)

def display_selected_image(image_path):
    # Load selected image with Pillow
    pil_image = Image.open(image_path)

    # Convert image to RGB format
    pil_image_rgb = pil_image.convert("RGB")

    # Resize the image to fit the canvas while maintaining aspect ratio
    max_width = 800  # Set your desired maximum width
    img_width, img_height = pil_image_rgb.size
    if img_width > max_width:
        ratio = max_width / img_width
        img_width = max_width
        img_height = int(img_height * ratio)

    pil_image_resized = pil_image_rgb.resize((img_width, img_height), Image.ANTIALIAS)

    # Display resized image using Tkinter Canvas
    img_tk = ImageTk.PhotoImage(pil_image_resized)
    canvas.config(width=img_width, height=img_height)
    canvas.create_image(0, 0, anchor='nw', image=img_tk)
    canvas.img_tk = img_tk  # Keep a reference to avoid garbage collection

def search_results():
    global output_list, seen_names, output_text

    search_term = entry_search.get().lower()
    if search_term:
        matching_results = set(result["Name"].lower() for result in output_list if search_term in result["Name"].lower())
        filtered_output_text = ""
        for output in output_list:
            if output["Name"].lower() in matching_results:
                filtered_output_text += "Name: {}\nHealth Point: {}\nWhat's the use?: {}\n\n".format(
                    output['Name'], output['Health Point'], output['What\'s the use?'])
        result_text.config(state="normal")
        result_text.delete(1.0, "end")
        result_text.insert("insert", filtered_output_text)
        result_text.config(state="disabled")
    else:
        result_text.config(state="normal")
        result_text.delete(1.0, "end")
        result_text.insert("insert", output_text)
        result_text.config(state="disabled")

# Tkinter GUI
root = Tk()
root.title("Image Processing with OCR")
root.geometry("900x600")  # Set the initial size of the window

# Frame for buttons and labels
button_frame = Frame(root, bg="#e6e6e6")  # Light gray background
button_frame.pack(pady=10)

open_button = Button(button_frame, text="Open Image", command=open_file_dialog, bg="#4caf50", fg="white")  # Green button
open_button.pack(side="left", padx=10)

path_label = Label(button_frame, text="Image Path: ", bg="#e6e6e6")  # Light gray background
path_label.pack(side="left")

process_button = Button(button_frame, text="Process Image", command=process_image, bg="#2196f3", fg="white")  # Blue button
process_button.pack(side="left", padx=10)

# Entry and Button for search
entry_search = Entry(button_frame, width=20)
entry_search.pack(side="right", padx=10)
search_button = Button(button_frame, text="Search", command=search_results, bg="#ff9800", fg="white")  # Orange button
search_button.pack(side="right", padx=10)

# Main frame for Canvas and Text
main_frame = Frame(root, bg="#ffffff")  # White background
main_frame.pack(expand=True, fill="both")

# Canvas for displaying the processed image
canvas = Canvas(main_frame, bg="#ffffff")  # White background
canvas.pack(side="left", padx=10, fill="both", expand=True)

# Scrollbar and Text widget for displaying output
scrollbar = Scrollbar(main_frame, orient="vertical")
scrollbar.pack(side="right", fill="y")

result_text = Text(main_frame, wrap="word", yscrollcommand=scrollbar.set, state="disabled", height=10, width=50, bg="#f2f2f2")  # Light gray background
result_text.pack(pady=10, padx=10, side="left", fill="both", expand=True)

scrollbar.config(command=result_text.yview)

root.mainloop()
