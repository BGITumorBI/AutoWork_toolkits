import os
from pptx import Presentation

def create_ppt_from_images(directory, output_file):
    # 创建 PowerPoint 对象
    prs = Presentation()

    # 遍历目录中的图片文件
    for filename in os.listdir(directory):
        if filename.endswith(".jpg") or filename.endswith(".png"):
            # 创建新的幻灯片
            slide_layout = prs.slide_layouts[1]  # 使用第二个默认布局（标题和内容）
            slide = prs.slides.add_slide(slide_layout)

            # 添加图片到幻灯片
            image_path = os.path.join(directory, filename)
            slide.shapes.add_picture(image_path, 0, 0, width=prs.slide_width, height=prs.slide_height)

    # 保存 PPT 文件
    prs.save(output_file)

# 指定目录和输出文件名
image_directory = "test"
output_ppt = "output.pptx"

# 创建 PPT
create_ppt_from_images(image_directory, output_ppt)
