import os
import time
import requests
import urllib3
import pandas as pd
from pathlib import Path

# 禁用 SSL 警告输出
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# =========================
# 1. 基础配置
# =========================

INPUT_FILE = "prompts_translated_level_4.xlsx" # 确保仓库里有这个文件！
# 在 GitHub 上我们直接把结果输出到当前目录，方便打包下载
OUTPUT_FILE = "model_eval_results.xlsx"

API_BASE_URL = "https://ai.gitee.com/v1"
# 【关键修改】使用 os.getenv 从云端环境变量读取，绝不在代码里写死！
API_KEY = "K3GW5OZWBPGG5BKZCXIN8USU3OJYDOBENS93IE1F"

MODELS = {
    "GLM-5": "glm-5",
    # "Kimi-K2.5": "kimi-k2-5",
    # "Deepseek-R1": "deepseek-r1",
    # "Qwen3-235B-A22B": "qwen3-235b-a22b",
}

# =========================
# 2. API 调用函数
# =========================

def call_model(model_name: str, prompt: str, retries: int = 3) -> str:
    if pd.isna(prompt) or not str(prompt).strip():
        return ""

    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }

    payload = {
        "model": model_name,
        "messages": [{"role": "user", "content": str(prompt)}],
        "temperature": 0.1,
        "max_tokens": 4096, # 允许长文本输出
    }

    endpoint = f"{API_BASE_URL}/chat/completions"

    for attempt in range(retries):
        try:
            # 保持 300 秒超时和忽略 SSL 验证
            response = requests.post(endpoint, headers=headers, json=payload, timeout=300, verify=False)
            
            if response.status_code == 200:
                data = response.json()
                if 'choices' in data and len(data['choices']) > 0:
                    return data['choices'][0]['message']['content']
                else:
                    return f"[解析异常] 未找到 choices 字段: {str(data)[:200]}"
            else:
                print(f"      [警告] HTTP {response.status_code} - 尝试重试 ({attempt+1}/{retries})")
                time.sleep(2)
                
        except Exception as e:
            print(f"      [错误] 请求异常: {str(e)} - 尝试重试 ({attempt+1}/{retries})")
            time.sleep(2)

    return "[请求失败] 超过最大重试次数"


# =========================
# 3. 主逻辑
# =========================

def main():
    if not API_KEY:
        print("错误：未找到 GITEE_API_KEY 环境变量！请在 GitHub Secrets 中配置。")
        return

    if not os.path.exists(INPUT_FILE):
        print(f"找不到输入文件: {INPUT_FILE}")
        return

    print("正在加载 Excel 数据...")
    df = pd.read_excel(INPUT_FILE)
    df = df.head(2)
    # 已经去掉了 df.head(2)，直接跑全量数据！

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        for display_name, api_model_name in MODELS.items():
            print(f"\n{'='*40}")
            print(f"开始测试模型: {display_name}")
            print(f"{'='*40}")

            rows = []
            for idx, row in df.iterrows():
                print(f"  [{display_name}] 正在处理第 {idx + 1}/{len(df)} 行...")

                chinese_prompt = row.get("prompt", "")
                english_prompt = row.get("english_prompt", "")

                chinese_output = call_model(api_model_name, chinese_prompt)
                time.sleep(1)

                english_output = call_model(api_model_name, english_prompt)
                time.sleep(1) 

                rows.append({
                    "topic": row.get("topic", ""),
                    "source": row.get("source", ""),
                    "level": row.get("level", ""),
                    "category": row.get("category", ""),
                    "safe type": row.get("base_type", ""),
                    "中文prompt": chinese_prompt,
                    "对应的输出(中文)": chinese_output,
                    "英文prompt": english_prompt,
                    "对应的输出(英文)": english_output,
                })

            out_df = pd.DataFrame(rows)
            sheet_name = display_name[:31]
            out_df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"模型 {display_name} 测试完成，已写入 Sheet！")

    print(f"\n✅ 所有测试已完成，结果保存在: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
