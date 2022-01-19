import cv2
import numpy as np
import matplotlib.pyplot as plt


# image path:
#path = "D://opencvImages//"
#fileName = "out.jpg"

# Reading an image in default mode:
for i in range(111,125):
    if i == 120 or i == 114 or i == 115 or i == 119:
        continue
    else:
        w = 'data/pics5/%d_5.png' %(i)
        inputImage = cv2.imread(w)

        # Convert RGB to grayscale:
        grayscaleImage = cv2.cvtColor(inputImage, cv2.COLOR_BGR2GRAY)

        # Convert the BGR image to HSV:
        hsvImage = cv2.cvtColor(inputImage, cv2.COLOR_BGR2HSV)

        # Create the HSV range for the blue ink:
        # [128, 255, 255], [90, 50, 70]
        lowerValues = np.array([64, 64, 64])
        upperValues = np.array([164, 264, 264])

        # Get binary mask of the blue ink:
        bluepenMask = cv2.inRange(hsvImage, lowerValues, upperValues)
        m = cv2.bitwise_not(bluepenMask)
        cv2.imwrite('data2/pics5/%d_5.png'%(i),m)







# Use a little bit of morphology to clean the mask:
# Set kernel (structuring element) size:
##kernelSize = 3
# Set morph operation iterations:
##opIterations = 1
# Get the structuring element:
##morphKernel = cv2.getStructuringElement(cv2.MORPH_RECT, (kernelSize, kernelSize))
# Perform closing:
##bluepenMask = cv2.morphologyEx(bluepenMask, cv2.MORPH_CLOSE, morphKernel, None, None, opIterations, cv2.BORDER_REFLECT101)

# Add the white mask to the grayscale image:
##colorMask = cv2.add(grayscaleImage, bluepenMask)
##_, binaryImage = cv2.threshold(colorMask, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
##cv2.imwrite('data2/bwimage.png',binaryImage)
##thresh, im_bw = cv2.threshold(binaryImage, 210, 230, cv2.THRESH_BINARY)
##kernel = np.ones((1, 1), np.uint8)
##imgfinal = cv2.dilate(im_bw, kernel=kernel, iterations=1)

