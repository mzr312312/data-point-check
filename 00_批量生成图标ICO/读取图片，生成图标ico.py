from PIL import Image

# 打开原始图像
image = Image.open("对钩.png")

# 定义需要的分辨率
sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]

# 保存为 .ico 文件
image.save("对钩.ico", format="ICO", sizes=sizes)