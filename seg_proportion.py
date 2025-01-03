import os
import json
import openpyxl
from openpyxl import Workbook
from shapely.geometry import Polygon

type = 1
def calculate_area(points):
    """Calculate the area of the polygon given its points."""
    polygon = Polygon(points)
    return polygon.area


def process_json(file_path, image_area):
    with open(file_path, 'r') as file:
        data = json.load(file)

    labels_info = {}
    total_label_area = 0

    for shape in data['shapes']:
        label = shape['label']
        points = shape['points']
        area = calculate_area(points)
        total_label_area += area
        if label in labels_info:
            labels_info[label] += area
        else:
            labels_info[label] = area

    for label in labels_info:
        area = labels_info[label]
        proportion = area / image_area
        labels_info[label] = f"{int(area)}/{int(image_area)} = {proportion:.5%}" if type == 0 else f"{proportion:.5%}"

    background_area = image_area - total_label_area
    background_proportion = background_area / image_area
    labels_info['background'] = f"{int(background_area)}/{int(image_area)} = {background_proportion:.5%}" if type == 0 else f"{background_proportion:.5%}"

    return labels_info


def process_directory(image_dir, json_dir, output_excel):
    all_labels = set()
    image_infos = []

    for image_filename in os.listdir(image_dir):
        if image_filename.lower().endswith(('.png', '.jpg', '.jpeg')):
            json_filename = os.path.splitext(image_filename)[0] + '.json'
            json_path = os.path.join(json_dir, json_filename)

            image_path = os.path.join(image_dir, image_filename)
            image = openpyxl.drawing.image.Image(image_path)
            image_width, image_height = image.width, image.height
            image_area = image_width * image_height

            if os.path.exists(json_path):
                labels_info = process_json(json_path, image_area)
            else:
                labels_info = {'background': f"{int(image_area)}/{int(image_area)} = 100.00%"} if type == 0 else {'background': "100.00%"}

            image_infos.append((image_filename, labels_info))
            all_labels.update(labels_info.keys())

    all_labels = sorted(all_labels)
    if 'background' in all_labels:
        all_labels.remove('background')
    all_labels.insert(0, 'background')

    output_dir = os.path.dirname(output_excel)
    if not os.path.exists(output_dir) and output_dir != '':
        os.makedirs(output_dir)

    workbook = Workbook()
    sheet = workbook.active

    header = ['File Name'] + all_labels
    sheet.append(header)

    for filename, labels_info in image_infos:
        row = [filename]
        for label in all_labels:
            row.append(labels_info.get(label, '0/0 = 0.00%' if type == 0 else "0.00%"))
        sheet.append(row)

    workbook.save(output_excel)


# 使用方法
image_dir = r'F:\All_image\Beijing_all'  # 替换为你的图片文件夹路径
json_dir = r'F:\labelme\Beijing\Pipe'  # 替换为你的JSON文件夹路径
output_excel = r'F:\labelme\Pipe_Beijing_output.xlsx'  # 输出Excel文件的路径
process_directory(image_dir, json_dir, output_excel)
