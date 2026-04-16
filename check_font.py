import os
import glob

paths_to_check = [
    '/usr/share/fonts/noto/NotoSans-Regular.ttf',
    '/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf',
]

for p in paths_to_check:
    print(f"{p}: {os.path.exists(p)}")

# Also look for any NotoSans
fonts = glob.glob('/usr/share/fonts/**/NotoSans-Regular.ttf', recursive=True)
print("Found:", fonts)
