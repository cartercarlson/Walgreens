import os
import shutil
import cv2

def images_match(file1, file2):
    # Open image
    image1 = cv2.imread(dir + file1)
    image2 = cv2.imread(dir + file2)

    # End if image dimensions are different
    if image1.shape != image2.shape:
        return

    # Convert to grayscale
    image1 = cv2.cvtColor(image1, cv2.COLOR_BGR2GRAY)
    image2 = cv2.cvtColor(image2, cv2.COLOR_BGR2GRAY)

    # Standardize image size
    max_size = 100, 100
    image1 = cv2.resize(image1, max_size)
    image2 = cv2.resize(image2, max_size)

    # Compare values of pixels
    difference = cv2.subtract(image1, image2)

    # Give 10px margin of error in case image is slightly tinted
    pix_match = (difference >= 0).sum()
    percent_similar = pix_match / len(difference)

    if percent_similar > 0.90:
        return True

def missing_border(image):

    v_border = True
    h_border = True
    pix_req = 15

    # Grab dimensions
    height = image.shape[0]
    width = image.shape[1]

    # Loop over image
    # White rgb = [255, 255, 255] or 765 - give 10 px of leeway = 755
    for a in range(1, pix_req, 2):

        # Check corner borders
        if (image[a, a].sum() or
            image[width - a, a].sum() or
            image[a, height - a].sum() or
            image[width - a, height - a].sum()) < 755:
            return True

        # Check vertical borders
        if v_border:
            for b in range(pix_req, width - pix_req, 10):
                if (image[b, a].sum() or image[b, height - a].sum()) < 755:
                    v_border = False
                    break

        # Check horizontal borders
        if h_border:
            for b in range(pix_req, height - pix_req, 10):
                if(image[a, b].sum() or image[width - a, b].sum()) < 755:
                    h_border = False
                    break

    if not v_border and h_border:
        return True

shared_path = '//hqfas322003c.corp.drugstore.com/cnc1shared/Files to Clean/Walgreens.com/'
dir = shared_path + 'Images to process/'
dir_border = shared_path + 'REJECTED/Bad border/'
dir_duplicate = shared_path + 'REJECTED/Duplicate images/'

borders = []
duplicates = []
files = []

[files.append(file) for file in os.listdir(dir)]

# Compare duplicate images
for file in range(0, len(files) - 1):
    if files[file].endswith('db'):
        continue

    file1 = files[file]
    # Grab upc for file 1
    if '_' in files[file]:
        file1_upc = files[file][:files[file].find('_')]
    else:
        file1_upc = files[file][:files[file].find('.')]
    next_image = 1
    while True:
        try:
            file2 = files[file + next_image]
            # Grab upc for file 2
            if '_' in files[file + next_image]:
                file2_upc = files[file + next_image][:files[file + next_image].find('_')]
            else:
                file2_upc = files[file + next_image][:files[file + next_image].find('.')]
            # See if the upc's match
            if file1_upc != file2_upc:
                break
            elif images_match(file1, file2):
                duplicates.append(file1)
                duplicates.append(file2)
        except:
            break
        next_image += 1

# Keep one copy of each image file
duplicates = set(duplicates)

# Remove duplicate images so we don't test them for border
files = set(files) - duplicates
files = list(files)

# Check image for white border
for file in range(0, len(files)):
    if files[file].endswith('db'):
        continue

    image = cv2.imread(dir + files[file])

    if missing_border(image):
        borders.append(files[file])

# Move files
for file in borders:
    shutil.move(dir + file, dir_border + file)

for file in duplicates:
    shutil.move(dir + file, dir_duplicate + file)
