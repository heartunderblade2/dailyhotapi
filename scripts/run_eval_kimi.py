import os
import time
import requests
import pandas as pd
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

# =========================
# 1. 基础配置
# =========================

INPUT_FILE = "prompts_259.xlsx"

OUTPUT_DIR = Path("outputs")
OUTPUT_DIR.mkdir(exist_ok=True)

OUTPUT_FILE = OUTPUT_DIR / "result_kimi.xlsx"

API_BASE_URL = "https://api.moark.com/v1"

API_KEY = os.getenv("API_KEY")

if not API_KEY:
    raise ValueError("未设置 API_KEY 环境变量")

MODELS = {
    # "Deepseek-R1": "deepseek-r1",
    # "GLM-5": "glm-5",
    "Kimi-K2.5": "Kimi-K2.5",
    # "Qwen3-235B-A22B": "qwen3-235b-a22b",
    # "GPT-4o-mini": "openai/gpt-4o-mini",
    # "Gemini-2.5-Flash-Lite": "google/gemini-2.5-flash-lite",
}

# 并发数
# 可以先设 10，稳定后改成 20、30、50
MAX_WORKERS = 30


# =========================
# 2. API 调用函数
# =========================

def call_model(model_name: str, prompt: str, retries: int = 2) -> str:
    """
    调用模型 API
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

            if response.status_code == 200:

                data = response.json()

                if (
                    "choices" in data
                    and len(data["choices"]) > 0
                ):
                    return data["choices"][0]["message"]["content"]

                return "[解析异常] 未找到 choices"

            else:
                print(
                    f"[{model_name}] HTTP {response.status_code}"
                )

                try:
                    print(response.json())
                except Exception:
                    print(response.text)

                if response.status_code == 400:
                    return f"[HTTP 400] {response.text}"

                time.sleep(2)

        except Exception as e:

            print(f"[异常] {model_name}: {e}")

            time.sleep(2)

    return "[请求失败] 超过最大重试次数"


# =========================
# 3. 单行任务函数
# =========================

def process_one_row(idx, row, api_model_name):
    """
    并发执行的最小任务：
    一行里面包含中文 prompt 和英文 prompt。
    """

    chinese_prompt = row.get("prompt", "")
    english_prompt = row.get("english_prompt", "")

    chinese_output = call_model(
        api_model_name,
        chinese_prompt
    )

    # 如果你想更猛一点，可以删掉这个 sleep
    time.sleep(0.5)

    english_output = call_model(
        api_model_name,
        english_prompt
    )

    time.sleep(0.5)

    return idx, {
        "topic": row.get("topic", ""),
        "source": row.get("source", ""),
        "level": row.get("level", ""),
        "category": row.get("category", ""),
        "safe type": row.get("base_type", ""),

        "中文prompt": chinese_prompt,
        "对应的输出(中文)": chinese_output,

        "英文prompt": english_prompt,
        "对应的输出(英文)": english_output,
    }


# =========================
# 4. 主逻辑
# =========================

def main():

    if not os.path.exists(INPUT_FILE):
        print(f"找不到输入文件: {INPUT_FILE}")
        return

    print("正在加载 Excel 数据...")

    df = pd.read_excel(INPUT_FILE)

    # 测试时只跑前几行
    df = df.head(15)

    with pd.ExcelWriter(
        OUTPUT_FILE,
        engine="openpyxl"
    ) as writer:

        for display_name, api_model_name in MODELS.items():

            print(f"\n========== {display_name} ==========")

            results = {}

            with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:

                futures = []

                for idx, row in df.iterrows():

                    future = executor.submit(
                        process_one_row,
                        idx,
                        row,
                        api_model_name
                    )

                    futures.append(future)

                total = len(futures)
                finished = 0

                for future in as_completed(futures):

                    finished += 1

                    try:
                        idx, result = future.result()
                        results[idx] = result

                    except Exception as e:
                        print(f"[任务异常] {display_name}: {e}")

                    print(
                        f"[{display_name}] "
                        f"{finished}/{total}"
                    )

            # 按原始 Excel 顺序重新排列
            rows = []

            for idx in df.index:
                if idx in results:
                    rows.append(results[idx])
                else:
                    row = df.loc[idx]
                    rows.append({
                        "topic": row.get("topic", ""),
                        "source": row.get("source", ""),
                        "level": row.get("level", ""),
                        "category": row.get("category", ""),
                        "safe type": row.get("base_type", ""),

                        "中文prompt": row.get("prompt", ""),
                        "对应的输出(中文)": "[任务失败]",

                        "英文prompt": row.get("english_prompt", ""),
                        "对应的输出(英文)": "[任务失败]",
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
# 5. 入口
# =========================

if __name__ == "__main__":
    main()
