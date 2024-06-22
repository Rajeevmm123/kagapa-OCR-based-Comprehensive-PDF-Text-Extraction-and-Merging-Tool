import os
import cv2
import pytesseract
import fitz
import PySimpleGUI as sg

def convert_pdf_to_images(pdf_path, output_folder, page_start=0, page_end=None):
    pdf_document = fitz.open(pdf_path)
    
    # Ensure the output folder exists
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    image_paths = []

    # Iterate through pages in the PDF
    for page_num in range(page_start, min(page_end, pdf_document.page_count) if page_end else pdf_document.page_count):
        page = pdf_document[page_num]
        
        # Render the page as an image
        image = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72))
        image_path = os.path.join(output_folder, f"page_{page_num + 1}.png")
        
        # Save the image
        image.save(image_path, "png")
        image_paths.append(image_path)

    pdf_document.close()
    return image_paths

def process_images(images, output_folder, window):
    # Ensure the output folder exists
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Create a progress bar
    progress_bar = window['progressbar']

    # Iterate through images
    for i, image_path in enumerate(images):
        # Read the image
        image = cv2.imread(image_path)

        # Apply image processing techniques (e.g., resizing, cropping, enhancing)
        # Replace this with your specific processing steps based on the characteristics of your images

        # Perform OCR for text extraction
        text = pytesseract.image_to_string(image, lang="eng+kan+hin+tel+tam+san")

        # Save extracted text to a text file
        text_file_path = os.path.join(output_folder, f"{os.path.splitext(os.path.basename(image_path))[0]}.rtf")
        with open(text_file_path, 'w', encoding='utf-8') as text_file:
            text_file.write(text)

        # Update progress bar
        progress_bar.update_bar(i + 1, len(images))

    print("Image processing and cataloging completed.")

def merge_text_files(txt_folder, output_file):
    # Get all text files in the specified folder
    text_files = [f for f in os.listdir(txt_folder) if f.endswith('.rtf')]

    # Ensure the output folder exists
    if not os.path.exists(txt_folder):
        os.makedirs(txt_folder)

    # Merge all text files into one
    with open(output_file, 'w', encoding='utf-8') as merged_file:
        for text_file in text_files:
            with open(os.path.join(txt_folder, text_file), 'r', encoding='utf-8') as f:
                # Add a page break before each text file content
                merged_file.write(f"\nPage Break: {text_file}\n\n")
                merged_file.write(f.read())
                merged_file.write('\n')  # Add a newline between each file

    print(f"Text files merged into {output_file}")

def main():
    sg.theme('SystemDefault')  # Use the default system theme

    # Define the layout of the GUI
    layout = [
        [sg.Text("Select a PDF file:")],
        [sg.InputText(key='pdf_path', size=(40, 1)), sg.FileBrowse(file_types=(("PDF Files", "*.pdf"),))],
        [sg.Text("Select an output folder:")],
        [sg.InputText(key='output_folder', size=(40, 1)), sg.FolderBrowse()],
        [sg.Text("Starting Page:"), sg.InputText(key='page_start', size=(5, 1)), sg.Text("Ending Page:"), sg.InputText(key='page_end', size=(5, 1))],
        [sg.Button("Extract Text"), sg.Button("Merge Text Files"), sg.Button("Exit")],
        [sg.ProgressBar(1, orientation='h', size=(20, 20), key='progressbar')],
    ]

    # Create the window
    window = sg.Window('ಆಕಾರ-PDF to Text Extractor', layout, resizable=True)
    # Set the window icon
    window.set_icon(r'D:\CoverPages\aakarasample.ico')

    # Event loop
    while True:
        event, values = window.read()

        if event == sg.WIN_CLOSED or event == 'Exit':
            break
        elif event == 'Extract Text':
            pdf_path = values['pdf_path']
            output_folder = values['output_folder']
            page_start = int(values['page_start']) - 1 if values['page_start'] else 0
            page_end = int(values['page_end']) if values['page_end'] else None

            if pdf_path and output_folder:
                try:
                    # Convert PDF pages to images
                    image_paths = convert_pdf_to_images(pdf_path, output_folder, page_start, page_end)

                    # Process the images using Tesseract OCR
                    process_images(image_paths, output_folder, window)

                    sg.popup("Text extraction completed!", title="Success")
                except Exception as e:
                    sg.popup_error(f"Error: {str(e)}", title="Error")

        elif event == 'Merge Text Files':
            txt_folder = values['output_folder']
            output_file = os.path.join(txt_folder, 'merged_text.rtf')
            
            if txt_folder:
                try:
                    # Merge all text files into one with page breaks
                    merge_text_files(txt_folder, output_file)

                    sg.popup("Text files merged successfully!", title="Success")
                except Exception as e:
                    sg.popup_error(f"Error: {str(e)}", title="Error")

    window.close()

if __name__ == "__main__":
    main()
