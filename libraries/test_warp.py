import cv2
import numpy as np

file_path = 'data/image/page0.jpg'
base_img = "data/base_image/base.jpg"
per = 25
imgQ = cv2.imread(base_img)
h, w, c = imgQ.shape
orb = cv2.ORB_create(1000)
kp1, des1 = orb.detectAndCompute(imgQ, None)

img = cv2.imread(file_path)
kp2, des2 = orb.detectAndCompute(img, None)
bf = cv2.BFMatcher(cv2.NORM_HAMMING)
matches = bf.match(des2, des1)
match_sorted = sorted(matches, key=lambda x: x.distance)
good = match_sorted[:int(len(match_sorted) * (per / 100))]
srcPoints = np.float32([kp2[m.queryIdx].pt for m in good]).reshape(-1, 1, 2)
dstPoints = np.float32([kp1[m.trainIdx].pt for m in good]).reshape(-1, 1, 2)
M, _ = cv2.findHomography(srcPoints, dstPoints, cv2.RANSAC, 5.0)
imgScan = cv2.warpPerspective(img, M, (w, h))
cv2.imshow('wapr img',imgScan)
cv2.waitKey(0)