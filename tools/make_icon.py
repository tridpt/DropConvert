from pathlib import Path

from PIL import Image, ImageDraw


root = Path(__file__).resolve().parent.parent
destination = root / "assets" / "icon.ico"
destination.parent.mkdir(parents=True, exist_ok=True)
image = Image.new("RGBA", (256, 256), (0, 0, 0, 0))
draw = ImageDraw.Draw(image)
draw.rounded_rectangle((8, 8, 248, 248), radius=62, fill="#6457e9")
draw.rounded_rectangle((69, 44, 183, 206), radius=13, outline="white", width=13)
draw.line((93, 97, 158, 97), fill="white", width=11)
draw.line((93, 128, 158, 128), fill="white", width=11)
draw.line((93, 159, 137, 159), fill="white", width=11)
image.save(destination, sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
print(destination)
