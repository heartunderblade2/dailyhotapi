import os
import sys
import subprocess
from pathlib import Path

# ==================== 配置区 ====================
OUTPUTS_DIR = Path("outputs")
JUDGE_MODEL = "deepseek-r1"

# 脚本名称配置
SCRIPT_GENERATE = "auto_generate_prompt.py" # 阶段 0
SCRIPT_EVAL = "auto_eval_model.py"         # 阶段 1
SCRIPT_ANNOTATE = "auto_annotation.py"     # 阶段 2
SCRIPT_ANALYSIS = "run_analysis.py"        # 阶段 3

EVAL_MODELS = [
    {"id": "glm-5", "name": "GLM-5"},
    {"id": "deepseek-r1", "name": "Deepseek-R1"},
    {"id": "Kimi-K2.5", "name": "Kimi-K2.5"},
    {"id": "qwen3-235b-a22b", "name": "Qwen"}
]
# =================================================

def run_phase_0_generate_prompts():
    print("\n" + "="*50)
    print("🚀 [阶段 0/3] 开始生成测试集 (prompts.xlsx)...")
    print("="*50)
    if not os.path.exists(SCRIPT_GENERATE):
        print(f"❌ 找不到生成脚本 {SCRIPT_GENERATE}，请检查。")
        sys.exit(1)
    
    p = subprocess.Popen([sys.executable, SCRIPT_GENERATE])
    p.wait()
    if p.returncode != 0:
        print("❌ 阶段 0 生成失败。")
        sys.exit(1)
    print("✅ 阶段 0 生成完成。")

def run_phase_1_evaluation():
    print("\n" + "="*50)
    print("🚀 [阶段 1/3] 开始模型回复生成...")
    print("="*50)
    processes = []
    for model in EVAL_MODELS:
        print(f"-> 启动测试: {model['name']}")
        cmd = [sys.executable, SCRIPT_EVAL, "--model_id", model["id"], "--display_name", model["name"]]
        p = subprocess.Popen(cmd)
        processes.append((model["name"], p))

    for name, p in processes:
        p.wait()
        if p.returncode != 0:
            print(f"⚠️ 警告: 模型 {name} 生成失败。")

def run_phase_2_annotation():
    print("\n" + "="*50)
    print(f"🚀 [阶段 2/3] 开始 AI 自动标注 (裁判: {JUDGE_MODEL})...")
    print("="*50)
    target_files = [f for f in OUTPUTS_DIR.glob("result_*.xlsx") if not f.name.endswith("_annotated.xlsx")]
    
    processes = []
    for file_path in target_files:
        print(f"-> 启动标注: {file_path.name}")
        cmd = [sys.executable, SCRIPT_ANNOTATE, "--input_file", file_path.name, "--model_id", JUDGE_MODEL]
        p = subprocess.Popen(cmd)
        processes.append((file_path.name, p))

    for name, p in processes:
        p.wait()

def run_phase_3_analysis():
    print("\n" + "="*50)
    print("🚀 [阶段 3/3] 开始执行数据分析与绘图...")
    print("="*50)
    p = subprocess.Popen([sys.executable, SCRIPT_ANALYSIS])
    p.wait()

def main():
    print("🌟 自动化全流程启动：Prompt生成 -> 模型测试 -> AI标注 -> 数据分析 🌟")
    
    run_phase_0_generate_prompts()
    run_phase_1_evaluation()
    run_phase_2_annotation()
    run_phase_3_analysis()
    
    print("\n🎉 全部流程执行结束！")

if __name__ == "__main__":
    main()