import os
import shutil
import cv2
from missing_border import missing_border

dir = 'F:/Old/images/'
dir_border = 'F:/Old/images/missing border/'

borders = []
files = []

[files.append(file) for file in os.listdir(dir)]

# Check image for white border
for file in range(0, len(files)):
    if files[file].endswith('db'):
        continue

    image = cv2.imread(dir + files[file])
    image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

    if missing_border(image):
        borders.append(files[file])

for file in borders:
    shutil.move(dir + file, dir_border + file)

