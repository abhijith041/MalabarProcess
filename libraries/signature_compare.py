import cv2
import numpy as np
import keras_ocr
from skimage.metrics import structural_similarity as ssim
from pdf2image import convert_from_path
from detect_box import detect_voucher


class SignatureComparison:
    '''
    Compare the two signatiure in the pdf image
    '''

    def __init__(self, filePath) -> None:
        '''
        :parm filePath: path the pdf file
        '''
        self.filePath = filePath
        self.pipeline = keras_ocr.pipeline.Pipeline()

    def image_from_pdf(self):
        '''
        conver the pdf to image using pdf2image library 
        the coverted images are then save to data/image folder
        '''

        # Store Pdf with convert_from_path function
        images = convert_from_path(self.filePath)

        for i in range(len(images)):

            # Save pages as images in the pdf
            images[i].save('libraries/data/image/'+'page' + str(i) + '.jpg', 'JPEG')

    def rezize_image(self):
        '''
        Resize the image to the base image using cv2 orb.
        : return: Return the resized image.
        '''
        file_path = 'libraries/data/image/'+'page' + '0' + '.jpg'
        base_img = "libraries/data/base_image/base.jpg"
        detect_voucher(file_path)
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
        srcPoints = np.float32(
            [kp2[m.queryIdx].pt for m in good]).reshape(-1, 1, 2)
        dstPoints = np.float32(
            [kp1[m.trainIdx].pt for m in good]).reshape(-1, 1, 2)
        M, _ = cv2.findHomography(srcPoints, dstPoints, cv2.RANSAC, 5.0)
        imgScan = cv2.warpPerspective(img, M, (w, h))
        # cv2.imshow('wapr img',imgScan)
        # cv2.waitKey(0)
        return imgScan

    def midpoint(self, x1, y1, x2, y2):
        '''
        Compute the midpoint of the input XY coordinates.
        : Return: x,y midpoint
        '''
        x_mid = int((x1 + x2) / 2)
        y_mid = int((y1 + y2) / 2)
        return (x_mid, y_mid)
    def remove_white_space(self, image):
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        blur = cv2.GaussianBlur(gray, (25,25), 0)
        thresh = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
        noise_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3,3))
        opening = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, noise_kernel, iterations=2)
        close_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (7,7))
        close = cv2.morphologyEx(opening, cv2.MORPH_CLOSE, close_kernel, iterations=3)
        # Find enclosing boundingbox and crop ROI\n”,
        coords = cv2.findNonZero(close)
        x,y,w,h = cv2.boundingRect(coords)
        return image[y:y+h, x:x+w]

    def mask_img(self, img):
        '''
        clean the background of the signature image image
        : return: cleaned image 
        '''
        prediction_groups = self.pipeline.recognize([img])

        mask = np.zeros(img.shape[:2], dtype="uint8")
        for box in prediction_groups[0]:
            if (len(box[0]) > 1):
                x0, y0 = box[1][0]
                x1, y1 = box[1][1]
                x2, y2 = box[1][2]
                x3, y3 = box[1][3]

                # x_mid0, y_mid0 = self.midpoint(x1, y1, x2, y2)
                # x_mid1, y_mi1 = self.midpoint(x0, y0, x3, y3)

                # thickness = int(math.sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2))

                # cv2.line(mask, (x_mid0, y_mid0), (x_mid1, y_mi1), 255, thickness)

                img = cv2.rectangle(img, (int(x0), int(y0)),
                                    (int(x2), int(y2)), (255, 255, 255), -1)

        return img

    def crop_img(self, img):
        '''
        Crop the resized image to get the signatures.
        : return: signature array
        '''
        recv_img = self.mask_img(img[279:365, 26:190].copy())
        imgFIlter = self.remove_white_space(recv_img)
        #cv2.imshow('recv_img', imgFIlter)
        ver_img = self.mask_img(img[275:365, 197:318].copy())
        imgFIlter2 = self.remove_white_space(ver_img)
        recv_img = imgFIlter
        ver_img = imgFIlter2
        #cv2.imshow('ver_img', imgFIlter2)
        #cv2.waitKey(0)
        h, w, c = recv_img.shape
        width, height = (w + int((w / 100) * 80), h + int((h / 100) * 80))
        recv_img = cv2.resize(recv_img, (width, height))
        h, w, c = ver_img.shape
        ver_img = cv2.resize(
            ver_img, (w + int((w / 100) * 80), h + int((h / 100) * 80)))
        recv_img = cv2.cvtColor(recv_img, cv2.COLOR_BGR2GRAY)
        ver_img = cv2.cvtColor(ver_img, cv2.COLOR_BGR2GRAY)
        recv_img = cv2.threshold(
            recv_img, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
        ver_img = cv2.threshold(
            ver_img, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
        #status = cv2.imwrite(
            #'D:/Quadance Tech files/doc-exraction/signature_verification/signature_verification_using_SSIM/data/cv/' + "receiver_sign.jpg", recv_img)

        #status = cv2.imwrite(
            #'D:/Quadance Tech files/doc-exraction/signature_verification/signature_verification_using_SSIM/data/cv/' + "verified_sign.jpg", ver_img)

        return [recv_img, ver_img]

    def verify_sign(self, img):
        '''
        compare the two image using structural_similarity
        : return: the similarity index range 0 to 1
        '''
        inp_img = img[0]
        anc_img = img[1]
        wrong_image = cv2.resize(inp_img.copy(), (100, 100))
        original_image = cv2.resize(anc_img.copy(), (100, 100))
        return ssim(original_image, wrong_image)

    def compute_similarity(self):
        '''
        compare the signature in pdf and return the similarity index
        : return: the similarity index range 0 to 1
        '''
        self.image_from_pdf()
        imgRe = self.rezize_image()
        imgCrop = self.crop_img(imgRe)
        si = self.verify_sign(imgCrop)
        return si
