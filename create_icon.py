# create_icon.py
from PIL import Image, ImageDraw, ImageFont
import os

# 创建256x256的图标
size = 256
img = Image.new('RGB', (size, size), color='#2E7D32')
draw = ImageDraw.Draw(img)

# 绘制简单的图标（示例：一个文档图标）
# 外框
draw.rectangle([50, 50, 206, 206], outline='white', width=5)
# 内部线条（表示文档）
draw.line([80, 100, 176, 100], fill='white', width=4)
draw.line([80, 130, 176, 130], fill='white', width=4)
draw.line([80, 160, 140, 160], fill='white', width=4)

# 保存为PNG
img.save('icon_temp.png')

# 注意：需要安装Pillow库：pip install Pillow
# 然后使用在线工具将PNG转换为ICO格式
