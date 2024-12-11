from fastapi import FastAPI, File, UploadFile, Form
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
import fitz  # PyMuPDF
import shutil
import os
from PIL import Image, ImageDraw
import comtypes.client  # For converting DOC and PPT to PDF (Windows only)

app = FastAPI()

# Directories for storing files
UPLOAD_FOLDER = "uploaded_files"
IMAGE_FOLDER = "pdf_images"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(IMAGE_FOLDER, exist_ok=True)

# Serve static files for uploaded images
app.mount("/static", StaticFiles(directory=IMAGE_FOLDER), name="static")


@app.get("/", response_class=HTMLResponse)
async def main():
    return """
    <html>
        <head>
            <title>File Upload Preview</title>
        </head>
        <body>
            <h1>Upload a File (JPG, PNG, PDF, DOC, PPT)</h1>
            <form action="/upload/" method="post" enctype="multipart/form-data">
                <input type="file" name="file" accept="image/png, image/jpeg, application/pdf, application/msword, application/vnd.ms-powerpoint">
                
                <label for="copy_color">Copy Color:</label>
                <select name="copy_color" id="copy_color">
                    <option value="black_and_white">Black & White</option>
                    <option value="color">Color</option>
                </select>
                
                <label for="orientation">Orientation:</label>
                <select name="orientation" id="orientation">
                    <option value="portrait">Portrait</option>
                    <option value="landscape">Landscape</option>
                </select>
                
                <label for="paper_punch">Paper Punch:</label>
                <select name="paper_punch" id="paper_punch">
                    <option value="no_hole">No Hole</option>
                    <option value="two_holes">Two Holes</option>
                    <option value="three_holes">Three Holes</option>
                </select>
                
                <label for="paper_binding">Paper Binding:</label>
                <select name="paper_binding" id="paper_binding">
                    <option value="no_staple">No Staple</option>
                    <option value="corner_staple">Corner Staple</option>
                </select>
                
                <button type="submit">Upload</button>
            </form>
        </body>
    </html>
    """


def convert_to_pdf(input_path: str, output_path: str):
    """Convert DOC/DOCX or PPT/PPTX to PDF using comtypes."""
    file_extension = input_path.split('.')[-1].lower()
    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)

    try:
        if file_extension in ["doc", "docx"]:
            word = comtypes.client.CreateObject("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(input_path)
            doc.SaveAs(output_path, FileFormat=17)  # 17 = PDF format
            doc.Close()
            word.Quit()
        elif file_extension in ["ppt", "pptx"]:
            powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
            powerpoint.Visible = 1  # Debug mode; change to False after testing
            presentation = powerpoint.Presentations.Open(input_path, WithWindow=False)
            presentation.SaveAs(output_path, 32)  # 32 = PDF format
            presentation.Close()
            powerpoint.Quit()
    except Exception as e:
        raise RuntimeError(f"Error converting file to PDF: {e}")

    return output_path


@app.post("/upload/")
async def upload_file(
    file: UploadFile = File(...),
    copy_color: str = Form(...),
    orientation: str = Form(...),
    paper_punch: str = Form(...),
    paper_binding: str = Form(...)
):
    file_extension = file.filename.split('.')[-1].lower()
    uploaded_file_path = os.path.join(UPLOAD_FOLDER, file.filename)

    # Save the uploaded file
    with open(uploaded_file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    # Handle file conversion to PDF if needed
    if file_extension in ["doc", "docx", "ppt", "pptx"]:
        pdf_path = os.path.join(
            UPLOAD_FOLDER, f"{os.path.splitext(file.filename)[0]}.pdf"
        )
        pdf_path = convert_to_pdf(uploaded_file_path, pdf_path)
    elif file_extension == "pdf":
        pdf_path = uploaded_file_path
    elif file_extension in ["jpg", "jpeg", "png"]:
        image_path = os.path.join(IMAGE_FOLDER, file.filename)
        shutil.copyfile(uploaded_file_path, image_path)
        converted_images = [process_image(image_path, copy_color, orientation, paper_punch, paper_binding)]
        return generate_html_response(converted_images)
    else:
        return HTMLResponse(content="<h1>Unsupported file type!</h1>")

    # Convert PDF to images
    pdf_doc = fitz.open(pdf_path)
    converted_images = []

    for page_num in range(len(pdf_doc)):
        page = pdf_doc[page_num]
        pix = page.get_pixmap(dpi=150)
        image_path = os.path.join(
            IMAGE_FOLDER, f"{os.path.basename(pdf_path).split('.')[0]}_page_{page_num + 1}.jpg"
        )
        pix.save(image_path)
        converted_images.append(process_image(image_path, copy_color, orientation, paper_punch, paper_binding))

    pdf_doc.close()
    return generate_html_response(converted_images)


from PIL import Image, ImageDraw

def process_image(image_path: str, copy_color: str, orientation: str, paper_punch: str, paper_binding: str) -> str:
    """Apply transformations to an image based on copy_color, orientation, paper_punch, and paper_binding."""
    with Image.open(image_path) as img:
        # Apply grayscale if black_and_white is selected
        if copy_color == "black_and_white":
            img = img.convert("L")

        # Apply rotation for landscape orientation
        if orientation == "landscape":
            img = img.rotate(90, expand=True)

        # Add visual indicators for paper punch
        draw = ImageDraw.Draw(img)
        width, height = img.size

        if paper_punch == "two_holes":
            draw.ellipse((10, height // 3 - 10, 30, height // 3 + 10), fill="grey")  # First hole
            draw.ellipse((10, 2 * height // 3 - 10, 30, 2 * height // 3 + 10), fill="grey")  # Second hole
        elif paper_punch == "three_holes":
            draw.ellipse((10, height // 4 - 10, 30, height // 4 + 10), fill="grey")  # Top hole
            draw.ellipse((10, height // 2 - 10, 30, height // 2 + 10), fill="grey")  # Middle hole
            draw.ellipse((10, 3 * height // 4 - 10, 30, 3 * height // 4 + 10), fill="grey")  # Bottom hole

        # Add staple mark in top left
        if paper_binding == "corner_staple":
            # Coordinates for a rotated rectangle (staple)
            staple_coords = [
                (35, 40),  # Top-left corner
                (75, 30),  # Top-right corner (slightly rotated)
                (77, 35),  # Bottom-right corner
                (37, 45),  # Bottom-left corner (slightly rotated)
            ]
            draw.polygon(staple_coords, fill="grey")

        # Save the transformed image
        transformed_image_path = os.path.join(
            IMAGE_FOLDER, f"processed_{os.path.basename(image_path)}"
        )
        img.save(transformed_image_path)
        return transformed_image_path

def generate_html_response(image_files: list) -> HTMLResponse:
    """Generate HTML response to display images in a slider."""
    slider_html = """
    <div id="slider" style="width: 600px; overflow: hidden; margin: auto;">
        <div id="slides" style="display: flex; transition: transform 0.5s;">
    """
    for image_file in image_files:
        slider_html += f'<img src="/static/{os.path.basename(image_file)}" style="width: 600px; height: auto;">'
    slider_html += """
        </div>
    </div>
    <button onclick="moveSlider(-1)">Previous</button>
    <button onclick="moveSlider(1)">Next</button>
    <script>
        let currentIndex = 0;
        function moveSlider(direction) {
            const slides = document.getElementById('slides');
            const totalSlides = slides.children.length;
            currentIndex = (currentIndex + direction + totalSlides) % totalSlides;
            slides.style.transform = `translateX(-${currentIndex * 600}px)`;
        }
    </script>
    """

    return HTMLResponse(content=f"""
    <html>
        <head>
            <title>File Preview</title>
        </head>
        <body>
            <h1>File Preview with Selected Options</h1>
            {slider_html}
            <p><a href="/">Go back to upload another file</a></p>
        </body>
    </html>
    """)

