import cairosvg
from PIL import Image
import io

# Convert SVG to PNG in memory
png_data = cairosvg.svg2png(url='icon.svg', output_width=256, output_height=256)

# Open PNG with Pillow
img = Image.open(io.BytesIO(png_data))

# Save as ICO
img.save('icon.ico', format='ICO', sizes=[(16,16), (32,32), (48,48), (64,64), (128,128), (256,256)])