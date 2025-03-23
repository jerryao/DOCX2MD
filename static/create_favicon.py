import os
from PIL import Image, ImageDraw

print("当前工作目录:", os.getcwd())

# 创建一个32x32的图像，带Alpha通道（RGBA）
img = Image.new('RGBA', (32, 32), color=(255, 255, 255, 0))
draw = ImageDraw.Draw(img)

# 绘制一个简单的字母"M"图标
draw.rectangle([0, 0, 31, 31], outline=(0, 123, 255), fill=(255, 255, 255, 230))
draw.line([(8, 8), (8, 24)], fill=(0, 123, 255), width=3)
draw.line([(24, 8), (24, 24)], fill=(0, 123, 255), width=3)
draw.line([(8, 8), (16, 16)], fill=(0, 123, 255), width=3)
draw.line([(24, 8), (16, 16)], fill=(0, 123, 255), width=3)

# 确保使用绝对路径保存文件
current_dir = os.path.dirname(os.path.abspath(__file__))
favicon_path = os.path.join(current_dir, 'favicon.ico')
print("正在创建favicon.ico文件:", favicon_path)

# 保存为favicon.ico
img.save(favicon_path, sizes=[(16, 16), (32, 32)])

print("Favicon.ico 已创建") 