import os
import time
import requests
import pandas as pd
from pathlib import Path

INPUT_FILE = "prompts_translated_level_4.xlsx"
OUTPUT_DIR = Path("outputs")
OUTPUT_DIR.mkdir(exist_ok=True)
OUTPUT_FILE = OUTPUT_DIR / "result_en.xlsx"

API_BASE_URL = "https://api.moark.com/v1"
API_KEY = "K3GW5OZWBPGG5BKZCXIN8USU3OJYDOBENS93IE1F"

if not API_KEY:
    raise ValueError("未设置 API_KEY 环境变量")

MODELS = {
    "GLM-5": "glm-5",
}

def call_model(model_name: str, prompt: str, retries: int = 2) -> str:
    # (API调用逻辑与 run_zh.py 保持完全一致)
    if pd.isna(prompt) or not str(prompt).strip():
        return ""
    headers = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}
    payload = {"model": model_name, "messages": [{"role": "user", "content": str(prompt)}], "temperature": 0.1, "max_tokens": 4096}
    endpoint = f"{API_BASE_URL}/chat/completions"

    for attempt in range(retries):
        try:
            response = requests.post(endpoint, headers=headers, json=payload, timeout=300)
            if response.status_code == 200:
                data = response.json()
                if "choices" in data and len(data["choices"]) > 0:
                    return data["choices"][0]["message"]["content"]
                return "[解析异常] 未找到 choices"
            else:
                if response.status_code == 400:
                    return f"[HTTP 400] {response.text}"
                time.sleep(2)
        except Exception as e:
            time.sleep(2)
    return "[请求失败] 超过最大重试次数"

def main():
    print("正在加载 Excel 数据...")
    df = pd.read_excel(INPUT_FILE)
    
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        for display_name, api_model_name in MODELS.items():
            print(f"\n========== {display_name} (英文端) ==========")
            rows = []
            
            for idx, row in df.iterrows():
                print(f"[{display_name}] 英文进度: {idx + 1}/{len(df)}")
                english_prompt = row.get("english_prompt", "")
                
                # 只请求英文
                english_output = call_model(api_model_name, english_prompt)
                time.sleep(1) # 防限流
                
                rows.append({
                    "topic": row.get("topic", ""),
                    "source": row.get("source", ""),
                    "level": row.get("level", ""),
                    "category": row.get("category", ""),
                    "safe type": row.get("base_type", ""),
                    "英文prompt": english_prompt,
                    "对应的输出(英文)": english_output,
                })
                
            out_df = pd.DataFrame(rows)
            out_df.to_excel(writer, sheet_name=display_name[:31], index=False)

    print(f"\n✅ 英文端全部完成，结果已保存至: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()