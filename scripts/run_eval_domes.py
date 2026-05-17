import os
import time
import requests
import pandas as pd
from pathlib import Path

# =========================
# 配置
# =========================

INPUT_FILE = "prompts_translated_level_4.xlsx"
OUTPUT_DIR = Path("outputs")
OUTPUT_DIR.mkdir(exist_ok=True)

OUTPUT_FILE = OUTPUT_DIR / "model_eval_results.xlsx"

MODELS = {
    "GLM-5": "glm-5",
    "Kimi-K2.5": "kimi-k2-5",
    "Deepseek-R1": "deepseek-r1",
    "Qwen3-235B-A22B": "qwen3-235b-a22b",
}

API_BASE_URL = os.getenv("API_BASE_URL")
API_KEY = os.getenv("API_KEY")


# =========================
# 调用 API
# =========================

def call_model(model_name: str, prompt: str) -> str:
    """
    根据你的 API 结构修改这里
    """

    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json",
    }

    payload = {
        "model": model_name,
          "messages": [
            {
                "role": "user",
                "content": prompt,
            }
        ],
        "temperature": 0,
    }

    try:
        response = requests.post(
            API_BASE_URL,
            headers=headers,
            json=payload,
            timeout=300,
        )

        response.raise_for_status()

        data = response.json()

        # 根据你的 API 返回结构调整
        return data["choices"][0]["message"]["content"]

    except Exception as e:
        return f"ERROR: {str(e)}"


# =========================
# 主逻辑
# =========================

def main():
    df = pd.read_excel(INPUT_FILE)

    with pd.ExcelWriter(OUTPUT_FILE, engine="xlsxwriter") as writer:

        for display_name, api_model_name in MODELS.items():
            print(f"\n===== Running model: {display_name} =====")

            rows = []

            for idx, row in df.iterrows():
                print(f"[{display_name}] Processing row {idx + 1}/{len(df)}")

                chinese_prompt = str(row["prompt"])
                english_prompt = str(row["english_prompt"])

                chinese_output = call_model(api_model_name, chinese_prompt)

                time.sleep(1)

                english_output = call_model(api_model_name, english_prompt)
                rows.append({
                    "topic": row.get("topic", ""),
                    "source": row.get("source", ""),
                    "level": row.get("level", ""),
                    "category": row.get("category", ""),
                    "safe_type": row.get("base_type", ""),
                    "中文prompt": chinese_prompt,
                    "中文输出": chinese_output,
                    "英文prompt": english_prompt,
                    "英文输出": english_output,
                })

                time.sleep(1)

            out_df = pd.DataFrame(rows)

            # sheet name 最长 31 字符
            sheet_name = display_name[:31]

            out_df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"\nSaved results to: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
