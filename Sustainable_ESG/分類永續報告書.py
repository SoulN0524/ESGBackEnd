import os
import shutil

# 定義檔案分類函數
def classify_pdfs(folder_path):
    # 定義各個語言和版本的資料夾路徑
    chinese_folder = os.path.join(folder_path, "110永續報告書(中文)")
    english_folder = os.path.join(folder_path, "110永續報告書(英文)")
    revised_folder = os.path.join(folder_path, "110永續報告書(修訂版)")


    # 遍歷資料夾中的所有文件
    for file_name in os.listdir(folder_path):
        # 檢查是否為PDF文件
        if file_name.endswith(".pdf"):
            file_path = os.path.join(folder_path, file_name)
            # 根據檔案名稱中的關鍵字進行分類
            if "110_E" in file_name:  # 中英文分類的關鍵字 "CH"
                shutil.move(file_path, os.path.join(chinese_folder, file_name))
                print(f"已移動 {file_name} 到 Chinese 資料夾")
            elif "110_M" in file_name:  # 英文分類的關鍵字 "EN"
                shutil.move(file_path, os.path.join(english_folder, file_name))
                print(f"已移動 {file_name} 到 English 資料夾")
            elif "110" in file_name:  # 修訂版分類的關鍵字 "REV"
                shutil.move(file_path, os.path.join(revised_folder, file_name))
                print(f"已移動 {file_name} 到 Revised 資料夾")
            else:
                print(f"未能歸類的文件：{file_name}")

# 使用範例
if __name__ == "__main__":
    folder_path = r"C:\Users\lin78\OneDrive\文件\永續report"  # 替換為你的報告文件夾路徑
    classify_pdfs(folder_path)
