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
    # "Kimi-K2.5": "kimi-k2-5",
    # "Deepseek-R1": "deepseek-r1",
    # "Qwen3-235B-A22B": "qwen3-235b-a22b",
}

API_BASE_URL = os.getenv("API_BASE_URL")
API_KEY = os.getenv("API_KEY")

print("API_BASE_URL:", API_BASE_URL)
print("API_KEY exists:", API_KEY is not None)

# =========================
# 调用 API
# =========================

def call_model(model_name: str, prompt: str) -> str:

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
        "max_tokens": 512,
        "stream": False,
    }

    try:

        response = requests.post(
            API_BASE_URL,
            headers=headers,
            json=payload,
            timeout=300,
        )

        print("STATUS:", response.status_code)
        print("TEXT:", response.text[:2000])

        try:

            data = response.json()

            print("JSON:", data)

            return str(data)

        except Exception as je:

            return f"JSON ERROR: {je} | RAW: {response.text[:1000]}"

    except Exception as e:

        return f"REQUEST ERROR: {str(e)}"


# =========================
# 主逻辑
# =========================

def main():
    df = pd.read_excel(INPUT_FILE)

    df = df.head(1)

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
