from sewar.full_ref import mse, rmse, psnr, uqi, ssim, ergas, scc, rase, sam, msssim, vifp
from skimage.metrics import structural_similarity
import cv2 as cv
import time

start= time.time()
img1 = cv.imread(r"data2/pics1/116_1.png")
img2 = cv.imread(r"data2/pics2/116_2.png")
grayA = cv.cvtColor(img1, cv.COLOR_BGR2GRAY)
grayB = cv.cvtColor(img2, cv.COLOR_BGR2GRAY)

(score, diff) = structural_similarity(grayA, grayB, full=True)



print(mse(img1,img2))
print(rmse(img1,img2))
print(uqi(img1, img2))
print(ergas(img1, img2))
print(scc(img1, img2))
print(rase(img1, img2))
print(sam(img1, img2))
print(vifp(img1, img2))
print(score)

method = 'ORB'  # 'SIFT'
lowe_ratio = 0.89

if method == 'ORB':
    finder = cv.ORB_create()
elif method == 'SIFT':
    finder = cv.xfeatures2d.SIFT_create()

# find the keypoints and descriptors with SIFT
kp1, des1 = finder.detectAndCompute(img1, None)
kp2, des2 = finder.detectAndCompute(img2, None)

# BFMatcher with default params
bf = cv.BFMatcher()
matches = bf.knnMatch(des1, des2, k=2)

# Apply ratio test
good = []

for m, n in matches:
    if m.distance < lowe_ratio * n.distance:
        good.append([m])

print(len(good))
print((len(good)/500))
print(time.time()-start)