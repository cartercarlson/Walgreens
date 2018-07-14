import os
import shutil
import cv2
import cython
%load_ext cython

@cython.boundscheck(False)
cpdef unsigned char[:, :] missing_border(int T, unsigned char [:, :] image):
    # Set variable extension types
    cdef int a, b, width, height, pix_req
    cdef bint h_border, v_border

    v_border = True
    h_border = True
    pix_req = 15

    # Grab dimensions
    height = image.shape[0]
    width = image.shape[1]

    # Loop over image
    for a in range(1, pix_req, 2):

        # Check corner borders
        if (sum(image[a, a]) or
            sum(image[width - a, a]) or
            sum(image[a, height - a]) or
            sum(image[width - a, height - a])):
            return True

        # Check vertical borders
        if v_border:
            for b in range(pix_req, width - pix_req, 10):
                if (sum(image[b, a]) or sum(image[b, height - a])):
                    v_border = False
                    break

        # Check horizontal borders
        if h_border:
            for b in range(pix_req, height - pix_req, 10):
                if(sum(image[a, b]) or sum(image[width - a, b])):
                    h_border = False
                    break

    if not v_border and h_border:
        return True


dir = 'C:/Users/Carter Carlson/Documents/Images/test/'
dir_border = 'C:/Users/Carter Carlson/Documents/Images/border/'

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

for file in duplicates:
    shutil.move(dir + file, dir_duplicate + file)


