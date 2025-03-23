import os
import pandas as pd
from pydub import AudioSegment
from PIL import Image
from moviepy import VideoFileClip
from PyPDF2 import PdfReader
from docx import Document

def convert_audio(file_path, output_format):
    try:
        audio = AudioSegment.from_file(file_path)
        output_file = f"{os.path.splitext(file_path)[0]}.{output_format.lower()}"
        audio.export(output_file, format=output_format)
        print(f"Audio converted to {output_file}")
    except Exception as e:
        print(f"Error converting audio: {e}")

def convert_image(file_path, output_format):
    try:
        image = Image.open(file_path)
        output_file = f"{os.path.splitext(file_path)[0]}.{output_format.lower()}"
        image.save(output_file, format=output_format.upper())
        print(f"Image converted to {output_file}")
    except Exception as e:
        print(f"Error converting image: {e}")

def convert_video(file_path, output_format):
    try:
        video = VideoFileClip(file_path)
        output_file = f"{os.path.splitext(file_path)[0]}.{output_format.lower()}"
        video.write_videofile(output_file, codec='libx264')
        print(f"Video converted to {output_file}")
    except Exception as e:
        print(f"Error converting video: {e}")

def convert_csv_to_xlsx(file_path):
    try:
        df = pd.read_csv(file_path)
        output_file = f"{os.path.splitext(file_path)[0]}.xlsx"
        df.to_excel(output_file, index=False)
        print(f"CSV converted to {output_file}")
    except Exception as e:
        print(f"Error converting CSV to XLSX: {e}")

def convert_xlsx_to_csv(file_path):
    try:
        df = pd.read_excel(file_path)
        output_file = f"{os.path.splitext(file_path)[0]}.csv"
        df.to_csv(output_file, index=False)
        print(f"XLSX converted to {output_file}")
    except Exception as e:
        print(f"Error converting XLSX to CSV: {e}")

def extract_text_from_pdf(file_path):
    try:
        reader = PdfReader(file_path)
        text = ""
        for page in reader.pages:
            text += page.extract_text() + "\n"
        output_file = f"{os.path.splitext(file_path)[0]}.txt"
        with open(output_file, 'w') as f:
            f.write(text)
        print(f"Text extracted from PDF to {output_file}")
    except Exception as e:
        print(f"Error extracting text from PDF: {e}")

def convert_docx_to_txt(file_path):
    try:
        doc = Document(file_path)
        text = '\n'.join([para.text for para in doc.paragraphs])
        output_file = f"{os.path.splitext(file_path)[0]}.txt"
        with open(output_file, 'w') as f:
            f.write(text)
        print(f"DOCX converted to {output_file}")
    except Exception as e:
        print(f"Error converting DOCX to TXT: {e}")

def get_conversion_choice(prompt, options):
    print(prompt)
    for i, option in enumerate(options, start=1):
        print(f"{i}. {option}")
    choice = input("Enter your choice: ").strip()
    return choice

def main():
    while True:
        print("\nFile Converter")
        print("1. Convert Audio")
        print("2. Convert Image")
        print("3. Convert Video")
        print("4. Convert CSV to XLSX")
        print("5. Convert XLSX to CSV")
        print("6. Extract Text from PDF")
        print("7. Convert DOCX to TXT")
        print("8. Exit")

        choice = input("Select an option (1-8): ").strip()

        if choice == '8':
            print("Exiting the program.")
            break

        file_path = input("Enter the file path: ").strip()
        if not os.path.isfile(file_path):
            print("File does not exist. Please try again.")
            continue

        try:
            if choice == '1' and file_path.lower().endswith(('.mp3', '.wav', '.ogg', '.flac', '.aac')):
                audio_choice = get_conversion_choice("Choose the output format for audio:", ['mp3', 'wav', 'ogg', 'flac', 'aac'])
                formats = ['mp3', 'wav', 'ogg', 'flac', 'aac']
                convert_audio(file_path, formats[int(audio_choice) - 1])
            elif choice == '2' and file_path.lower().endswith(('.jpeg', '.jpg', '.png', '.bmp', '.gif', '.tiff')):
                image_choice = get_conversion_choice("Choose the output format for image:", ['jpeg', 'jpg', 'png', 'bmp', 'gif', 'tiff'])
                formats = ['JPEG', 'JPEG', 'PNG', 'BMP', 'GIF', 'TIFF']
                convert_image(file_path, formats[int(image_choice) - 1])
            elif choice == '3' and file_path.lower().endswith(('.mp4', '.avi', '.mov', '.mkv', '.wmv')):
                video_choice = get_conversion_choice("Choose the output format for video:", ['mp4', 'avi', 'mov', 'mkv', 'wmv'])
                formats = ['mp4', 'avi', 'mov', 'mkv', 'wmv']
                convert_video(file_path, formats[int(video_choice) - 1])
            elif choice == '4' and file_path.lower().endswith('.csv'):
                convert_csv_to_xlsx(file_path)
            elif choice == '5' and file_path.lower().endswith('.xlsx'):
                convert_xlsx_to_csv(file_path)
            elif choice == '6' and file_path.lower().endswith('.pdf'):
                extract_text_from_pdf(file_path)
            elif choice == '7' and file_path.lower().endswith('.docx'):
                convert_docx_to_txt(file_path)
            else:
                print("Invalid choice or unsupported file format.")
        except (ValueError, IndexError):
            print("Invalid input. Please try again.")

if __name__ == "__main__":
    main()
