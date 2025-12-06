import os
from PIL import Image, ImageDraw
import subprocess
import sys

def create_dummy_images(folder):
    if not os.path.exists(folder):
        os.makedirs(folder)
    
    # Create 3 dummy images with different dimensions
    dimensions = [(800, 600), (600, 800), (1000, 1000)]
    colors = ['red', 'green', 'blue']
    
    for i, (dim, color) in enumerate(zip(dimensions, colors)):
        img = Image.new('RGB', dim, color=color)
        d = ImageDraw.Draw(img)
        d.text((10, 10), f"Image {i+1}", fill="white")
        filename = os.path.join(folder, f"test_image_{i+1}.jpg")
        img.save(filename)
        print(f"Created {filename}")

def run_converter():
    print("Running converter...")
    result = subprocess.run([sys.executable, "images_to_word.py", "test_images", "test_output.docx"], capture_output=True, text=True)
    print(result.stdout)
    if result.stderr:
        print("Errors:", result.stderr)
    
    if os.path.exists("test_output.docx"):
        print("SUCCESS: test_output.docx created.")
    else:
        print("FAILURE: test_output.docx not found.")

if __name__ == "__main__":
    create_dummy_images("test_images")
    run_converter()
