import os
import time
import requests
import pandas as pd
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

# =========================
# 1. 基础配置 (保留原 OpenRouter 配置)
# =========================

INPUT_FILE = "prompts_259.xlsx"

OUTPUT_DIR = Path("outputs")
OUTPUT_DIR.mkdir(exist_ok=True)

OUTPUT_FILE = OUTPUT_DIR / "result_intern.xlsx"

# OpenRouter 配置
API_BASE_URL = "https://openrouter.ai/api/v1"

# 从环境变量读取 Key
API_KEY = os.getenv("OPENROUTER_API_KEY")

if not API_KEY:
    raise ValueError("未设置 OPENROUTER_API_KEY 环境变量")

# 需要测试的模型
MODELS = {
    "Gemini-3.1-Flash-Lite": "google/gemini-3.1-flash-lite",
    "GPT-5-mini": "openai/gpt-5-mini",
    # 可以继续添加
    # "DeepSeek-R1": "deepseek/deepseek-r1",
    # "Qwen3-235B": "qwen/qwen3-235b-a22b",
    # "Kimi-K2": "moonshotai/kimi-k2",
}

# 并发数
# 建议 OpenRouter 免费/低限额用户不要设置太高，10~20 是比较稳妥的值
MAX_WORKERS = 10


# =========================
# 2. API 调用函数
# =========================

def call_model(model_name: str, prompt: str, retries: int = 2) -> str:
    """
    调用 OpenRouter 模型
    """

    if pd.isna(prompt) or not str(prompt).strip():
        return ""

    # 保留原代码特有的 Headers
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json",
        # OpenRouter 推荐
        "HTTP-Referer": "https://github.com",
        "X-Title": "model-eval",
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
                except Exception:
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
# 3. 单行任务函数 (并发最小单元)
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

    # 稍微休眠防止触发并发速率限制 (Rate Limit)
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

    # 测试时只跑前几行，正式跑全量时请注释或删掉这行
    # df = df.head(10)

    with pd.ExcelWriter(
        OUTPUT_FILE,
        engine="openpyxl"
    ) as writer:
        for display_name, api_model_name in MODELS.items():
            print(f"\n========== {display_name} ==========")
            results = {}

            # 启动线程池执行并发
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

                # 实时监听完成状态
                for future in as_completed(futures):
                    finished += 1
                    try:
                        idx, result = future.result()
                        results[idx] = result
                    except Exception as e:
                        print(f"[任务异常] {display_name}: {e}")

                    # 打印进度
                    print(
                        f"[{display_name}] "
                        f"进度: {finished}/{total}"
                    )

            # 按原始 Excel 顺序重新排列 (因为并发返回的顺序是乱的)
            rows = []
            for idx in df.index:
                if idx in results:
                    rows.append(results[idx])
                else:
                    # 如果某一行因为不可知原因彻底挂掉，保留基础数据占位
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

            # Excel sheet 名称限制 31 个字符
            sheet_name = display_name[:31]
            out_df.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False
            )

            print(f"{display_name} 完成，已写入 Sheet: {sheet_name}")

    print(f"\n✅ 全部完成")
    print(f"结果文件: {OUTPUT_FILE}")


# =========================
# 5. 入口
# =========================

if __name__ == "__main__":
    main()
