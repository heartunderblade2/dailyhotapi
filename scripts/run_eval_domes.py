import os
import time
import requests
import pandas as pd
from pathlib import Path

# =========================
# 1. 基础配置
# =========================

INPUT_FILE = "prompts_translated_level_4.xlsx"

OUTPUT_DIR = Path("outputs")
OUTPUT_DIR.mkdir(exist_ok=True)

OUTPUT_FILE = OUTPUT_DIR / "model_eval_results_domes.xlsx"

# OpenRouter 配置
API_BASE_URL = "https://ai.gitee.com/v1"

# 从环境变量读取 Key
API_KEY = os.getenv("API_KEY")

if not API_KEY:
    raise ValueError("未设置 API_KEY 环境变量")

# 需要测试的模型
MODELS = {
    # "Gemini-2.5-Flash-Lite": "google/gemini-2.5-flash-lite",
    # "GPT-4o-mini": "openai/gpt-4o-mini",
    # 可以继续添加
    "GLM-5": "glm-5",
    "Kimi-K2.5": "Kimi-K2.5",
    "Deepseek-R1": "deepseek-r1",
    "Qwen3-235B-A22B": "qwen3-235b-a22b",
}

# =========================
# 2. API 调用函数
# =========================

def call_model(model_name: str, prompt: str, retries: int = 2) -> str:
    """
    调用 OpenRouter 模型
    """

    if pd.isna(prompt) or not str(prompt).strip():
        return ""

    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json",
    }

    payload = {
        "model": model_name,
        "messages": [
            {
                "role": "user",
                "content": str(prompt)
            }
        ],
        "temperature": 0.1,
        "max_tokens": 4096,
    }

    endpoint = f"{API_BASE_URL}/chat/completions"

    for attempt in range(retries):

        try:
            response = requests.post(
                endpoint,
                headers=headers,
                json=payload,
                timeout=300
            )

            # 成功
            if response.status_code == 200:

                data = response.json()

                if (
                    "choices" in data
                    and len(data["choices"]) > 0
                ):
                    return data["choices"][0]["message"]["content"]

                return "[解析异常] 未找到 choices"

            # 失败
            else:
                print(
                    f"[{model_name}] HTTP {response.status_code}"
                )

                try:
                    print(response.json())
                except:
                    print(response.text)

                # 400 不需要重试
                if response.status_code == 400:
                    return f"[HTTP 400] {response.text}"

                time.sleep(2)

        except Exception as e:

            print(f"[异常] {model_name}: {e}")

            time.sleep(2)

    return "[请求失败] 超过最大重试次数"


# =========================
# 3. 主逻辑
# =========================

def main():

    if not os.path.exists(INPUT_FILE):
        print(f"找不到输入文件: {INPUT_FILE}")
        return

    print("正在加载 Excel 数据...")

    df = pd.read_excel(INPUT_FILE)

    # 测试时只跑前几行
    # 正式跑全量时删掉这行
    df = df.head(2)

    with pd.ExcelWriter(
        OUTPUT_FILE,
        engine="openpyxl"
    ) as writer:

        for display_name, api_model_name in MODELS.items():

            print(f"\n========== {display_name} ==========")

            rows = []

            for idx, row in df.iterrows():

                print(
                    f"[{display_name}] "
                    f"{idx + 1}/{len(df)}"
                )

                chinese_prompt = row.get("prompt", "")
                english_prompt = row.get("english_prompt", "")

                # 中文
                chinese_output = call_model(
                    api_model_name,
                    chinese_prompt
                )

                time.sleep(1)

                # 英文
                english_output = call_model(
                    api_model_name,
                    english_prompt
                )

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

            out_df.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False
            )

            print(f"{display_name} 完成")

    print(f"\n✅ 全部完成")
    print(f"结果文件: {OUTPUT_FILE}")


# =========================
# 4. 入口
# =========================

if __name__ == "__main__":
    main()
