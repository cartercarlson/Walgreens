import os
import shutil
from PIL import Image

dir = '//hqfas322003c.corp.drugstore.com/cnc1shared/Files to Clean/Walgreens.com/Images to process/'
dir_border = '//hqfas322003c.corp.drugstore.com/cnc1shared/Files to Clean/Walgreens.com/Bad image border/'
bad_borders = []

for file in os.listdir(dir):
    if file.endswith('db'):
        continue
    im = Image.open(dir + file)
    pix = im.load()
    pic_size = im.size
    width = pic_size[0]
    height = pic_size[1]
    avg_border = (width + height)/2
    pix_req = int(round(avg_border * 0.016))
    mid_height = height - pix_req * 2
    white_pixels = []

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

    for a in range(1, int(pix_req/2)):
        # pix[width, height]
        for b in range(1, int(width/10)):
            # Top portion
            white_pixels.extend(pix[b*10, a*2])
            # Bottom portion
            white_pixels.extend(pix[b*10, height - pix_req + a*2])
        # Two middle portions - between top and bottom portion
        for b in range(1, int(mid_height/10)):
            # Left side
            white_pixels.extend(pix[a*2, pix_req - 1 + b*10])
            # Right side
            white_pixels.extend(pix[width - pix_req - 1 + 2*a , pix_req - 1 + b*10])

    if not (all(for i in white_pixels if i == 255)):
        bad_borders.append(file)

for file in bad_borders:
    shutil.move(dir + file, dir_border + file)
