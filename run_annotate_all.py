import os
import subprocess
from pathlib import Path

# ==================== 配置 ====================
OUTPUTS_DIR = Path("outputs")
# 裁判模型，这里统一用 deepseek-r1 跑所有文件的标注任务
JUDGE_MODEL = "deepseek-r1" 

def main():
    if not OUTPUTS_DIR.exists():
        print(f"❌ 找不到 {OUTPUTS_DIR} 文件夹，请先运行上一个脚本生成测试结果。")
        return

    # 1. 扫描 outputs 文件夹下所有的需要标注的 Excel 文件
    target_files = []
    for file in OUTPUTS_DIR.glob("*.xlsx"):
        # 匹配前缀 (假设之前生成的文件叫 result_xxx.xlsx) 并且过滤掉已经加了 _annotated 后缀的文件
        if file.name.startswith("result_") and not file.name.endswith("_annotated.xlsx"):
            target_files.append(file.name)

    if not target_files:
        print("✅ outputs 文件夹中没有找到需要标注的新文件。")
        return

    print(f"🚀 找到 {len(target_files)} 个待标注文件，开始并发标注任务...\n")

    # 2. 并发启动子进程
    processes = []
    for file_name in target_files:
        print(f"-> 正在启动针对 {file_name} 的标注任务 (使用裁判模型: {JUDGE_MODEL})")
        
        # 构造命令行参数
        # 注意：这里默认你使用的是 python3，如果你系统里是 python，请改成 "python"
        cmd = [
            "python", "auto_annotation.py", 
            "--input_file", file_name, 
            "--model_id", JUDGE_MODEL
        ]
        
        # Popen 在后台异步执行
        p = subprocess.Popen(cmd)
        processes.append(p)

    # 3. 等待所有后台任务完成
    for p in processes:
        p.wait()

    print("\n🎉 所有标注任务执行完毕！")
    print("请前往 outputs 文件夹查看带 _annotated.xlsx 后缀的结果文件。")

if __name__ == "__main__":
    main()