from PIL import Image, ImageDraw
img = Image.new('RGB', (300, 300), color='#999999')
draw = ImageDraw.Draw(img)
# Draw a thick black border
draw.rectangle([0, 0, 299, 299], outline='black', width=5)
img.save('assets/default.png')
print("High-contrast assets/default.png created.")
