import subprocess

# 定义要测试的模型列表
models = [
    {"id": "glm-5", "name": "GLM-5"},
    {"id": "deepseek-r1", "name": "Deepseek-R1"},
    {"id": "Kimi-K2.5", "name": "Kimi-K2.5"},
    {"id": "qwen3-235b-a22b", "name": "Qwen"}
]

processes = []

print("🚀 正在启动多模型并发测试...\n")

for model in models:
    print(f"-> 启动测试: {model['name']} ({model['id']})")
    # 这里假设你的主测试脚本名字是 auto_eval_model.py
    cmd = ["python", "auto_eval_model.py", "--model_id", model["id"], "--display_name", model["name"]]
    
    # 启动子进程，不阻塞主进程
    p = subprocess.Popen(cmd)
    processes.append(p)

# 等待所有子进程运行结束
for p in processes:
    p.wait()

print("\n✅ 所有模型测试完毕！请检查 outputs 文件夹。")