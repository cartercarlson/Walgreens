import os
import shutil
from PIL import Image

dir = '//hqfas322003c.corp.drugstore.com/cnc1shared/Files to Clean/Walgreens.com/Images to process/'
dir_border = '//hqfas322003c.corp.drugstore.com/cnc1shared/Files to Clean/Walgreens.com/REJECTED/Bad image border/'
dir_duplicate = '//hqfas322003c.corp.drugstore.com/cnc1shared/Files to Clean/Walgreens.com/REJECTED/Duplicate image/'

bad_borders = []
duplicate_images = []
files = []

[files.append(file) for file in os.listdir(dir)]

# Note: Do I want to move them into their own folder? I think so...
# Compare duplicate images
for file in range(0, len(files) - 1):
    if files[file].endswith('db'):
        continue

    file1 = files[file]
    if '_' in files[file]:
        file1_upc = files[file][:files[file].find('_')]
    else:
        file1_upc = files[file][:files[file].find('.')]
    next_image = 1
    while True:
        file2 = files[file + next_image]
        if '_' in files[file + next_image]:
            file2_upc = files[file + next_image][:files[file + next_image].find('_')]
        else:
            file2_upc = files[file + next_image][:files[file + next_image].find('.')]
        if file1_upc != file2_upc:
            break
        next_image += 1
        if compare_images(file1, file2):
            duplicate_images.append(file1)
            duplicate_images.append(file2)
        else:
            break

import cv2
duplicate_images = duplicate_images.nunique()
def compare_images(file1, file2):

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

    difference = cv2.subtract(image1, image2)
    pix_match = (difference == 0).sum()
    pix_no_match = (difference > 0).sum()

    percent_similar = pix_match / (pix_match + pix_no_match)

    if percent_similar > 0.90:
        return True

    '''
    How dimensions for white border are calculated
    ----------------------------------------------
    top
    width= 0 to length
    height= 0 to pix_req

    bottom
    width= 0 to length
    height= height - pixreq to height

    left
    width= 0 to pix_req
    height= pix_req + 1 to height - (pix_req + 1)

    right
    width= width - pix_req to width
    height= pix_req + 1 to height - (pix_req + 1)
    ----------------------------------------------
    Note:
        a. Average of 1.6% of border should be white space
        b. This tests the border without validating EVERY pixel
        c. For top and bottom, it gets every 10th wide pixel, and every other height pixel
        d. for left and right, it gets every 10th height pixel, and every other width pixel
    '''
# Note: remove files from files list that are duplicate images
for file in range(0, len(files)):
    if files[file].endswith('db'):
        continue

    # Note: Do I have to load the first image twice?  Once in compare_images, once here
    im1 = Image.open(dir + files[file])
    pix = im1.load()
    pic_size = im1.size
    width, height = pic_size
    avg_dimension = (width + height)/2
    pix_req = int(round(avg_dimension * 0.016))
    mid_height = height - pix_req * 2
    white_pixels = []

    for a in range(1, int(pix_req), 2):
        # pix[width, height]
        for b in range(1, int(width), 10):
            # Top portion
            white_pixels.extend(pix[b, a])
            # Bottom portion
            white_pixels.extend(pix[b, height - pix_req + a])
        # Two middle portions - between top and bottom portion
        for b in range(1, int(mid_height), 10):
            # Left side
            white_pixels.extend(pix[a, pix_req - 1 + b])
            # Right side
            white_pixels.extend(pix[width - pix_req - 1 + a, pix_req - 1 + b])

    if not(all(for i in white_pixels if i == 255)):
        bad_borders.append(files[file])

for file in bad_borders:
    shutil.move(dir + file, dir_border + file)

for file in duplicate_images:
    shutil.move(dir + file, dir_duplicate + file)



'''
def compare_images(file1, file2):

    # Open image
    image1 = Image.open(dir + file1)
    image2 = Image.open(dir + file2)

    # Get image dimensions
    pic_size1 = image1.pic_size
    pic_size2 = image2.pic_size

    # End image comparison if images have different dimensions
    if pic_size1 != pic_size2:
        return

    # Convert to grayscale
    image1 = image1.convert('LA')
    image2 = image2.convert('LA')

    # Standardize image size
    size = 100, 100
    image1 = image1.thumbnail(size)
    image2 = image2.thumbnail(size)


    if image1 = 95% of image2:
        return True
'''
