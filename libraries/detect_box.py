from transformers import DetrImageProcessor, DetrForObjectDetection
import torch
from PIL import Image
import requests
import numpy as np


def create_bounding_box(points):
    """Create a bounding box from a list of points."""
    points = np.array(points)
    bbox = np.array([points.min(axis=0), points.max(axis=0)])
    return bbox

def inflate_bounding_box(bbox, amount):
    """Inflate a bounding box by a given amount."""
    bbox_center = np.mean(bbox, axis=0)
    bbox_size = np.abs(bbox[1] - bbox[0])
    bbox_inflated_size = bbox_size + amount * 2
    bbox_inflated_min = bbox_center - bbox_inflated_size / 2
    bbox_inflated_max = bbox_center + bbox_inflated_size / 2
    bbox_inflated = np.array([bbox_inflated_min, bbox_inflated_max])
    return bbox_inflated

def detect_voucher(path):
    '''
    detect object in an image using model and crop the first object, save the croped image in same path
    :parm path: path of the image
    '''
    image = Image.open(path)

    processor = DetrImageProcessor.from_pretrained(
        "TahaDouaji/detr-doc-table-detection")
    model = DetrForObjectDetection.from_pretrained(
        "TahaDouaji/detr-doc-table-detection")
    boxes = []
    inputs = processor(images=image, return_tensors="pt")
    outputs = model(**inputs)

    # convert outputs (bounding boxes and class logits) to COCO API
    # let's only keep detections with score > 0.9
    target_sizes = torch.tensor([image.size[::-1]])
    results = processor.post_process_object_detection(
        outputs, target_sizes=target_sizes, threshold=0.9)[0]

    for score, label, box in zip(results["scores"], results["labels"], results["boxes"]):
        box = [round(i, 2) for i in box.tolist()]
        print(
            f"Detected {model.config.id2label[label.item()]} with confidence "
            f"{round(score.item(), 3)} at location {box}"
        )

        boxes = results["boxes"]
    if len(boxes) > 1:
        box = [round(i, 2) for i in boxes[0].tolist()]
        boxVect = [(box[0], box[1]), (box[2], box[3])]

        bbox = create_bounding_box(boxVect)
        big_bbox = inflate_bounding_box(bbox,50)
        (b, c), (l, m) = big_bbox.min(axis=0), big_bbox.max(axis=0)
        #(b, c), (l, m) = big_bbox.min_point, big_bbox.max_point
        sc_box = [b, c, l, m]
        imageCrop = image.crop(box=sc_box)
        imageCrop.save(path)
