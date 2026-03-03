import os
from PIL import Image

def optimize_images():
    figures_dir = 'figures'
    for filename in os.listdir(figures_dir):
        if filename.endswith('.png'):
            png_path = os.path.join(figures_dir, filename)
            jpg_filename = filename.rsplit('.', 1)[0] + '.jpg'
            jpg_path = os.path.join(figures_dir, jpg_filename)
            
            try:
                # Open image and convert to RGB (removing alpha channel for JPG)
                with Image.open(png_path) as img:
                    rgb_im = img.convert('RGB')
                    # Save as optimized JPG
                    rgb_im.save(jpg_path, 'JPEG', quality=85, optimize=True)
                    print(f"Optimized {filename} -> {jpg_filename}")
            except Exception as e:
                print(f"Error optimizing {filename}: {e}")

if __name__ == "__main__":
    optimize_images()
