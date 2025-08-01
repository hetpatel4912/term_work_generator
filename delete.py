import os

for i in range(10,3300):
    if os.path.exists(f"/home/Hetindex/mysite/index{i}.docx"):
        file_path = f"/home/Hetindex/mysite/index{i}.docx"
        os.remove(file_path)