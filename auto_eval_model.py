import os
import time
import argparse
import requests
import pandas as pd
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

# =========================
# 1. 基础配置
# =========================

INPUT_FILE = "prompts.xlsx"
OUTPUT_DIR = Path("outputs")
OUTPUT_DIR.mkdir(exist_ok=True)

# 直接写死 URL 和 API_KEY
API_BASE_URL = "https://api.moark.com/v1"
API_KEY = "WFNMCKWOYPLYRHURRDW3BFSYYGND2EV26VLGW95F"  # <--- 请在此处替换为你的 API_KEY

# 并发数
MAX_WORKERS = 5


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
                if "choices" in data and len(data["choices"]) > 0:
                    return data["choices"][0]["message"]["content"]
                return "[解析异常] 未找到 choices"

            else:
                print(f"[{model_name}] HTTP {response.status_code}")
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
    并发执行的最小任务
    """
    chinese_prompt = row.get("prompt", "")
    english_prompt = row.get("english_prompt", "")

    chinese_output = call_model(api_model_name, chinese_prompt)
    time.sleep(0.5)

    english_output = call_model(api_model_name, english_prompt)
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
    if not API_KEY or API_KEY == "在这里填入你的真实_API_KEY":
        raise ValueError("请先在代码第 18 行填写你的 API_KEY。")

    # -------------------------
    # 解析命令行参数
    # -------------------------
    parser = argparse.ArgumentParser(description="单模型 API 并发测试脚本")
    parser.add_argument("--model_id", type=str, required=True, help="API 中的模型名称，如: glm-5, deepseek-r1")
    parser.add_argument("--display_name", type=str, default=None, help="结果文件中的展示名称，如: GLM-5")
    
    args = parser.parse_args()

    api_model_name = args.model_id
    display_name = args.display_name if args.display_name else api_model_name
    
    # 动态指定输出文件，防止并发脚本互相覆盖
    output_file = OUTPUT_DIR / f"result_{display_name.replace('/', '_')}.xlsx"

    if not os.path.exists(INPUT_FILE):
        print(f"找不到输入文件: {INPUT_FILE}")
        return

    print(f"正在加载 Excel 数据: {INPUT_FILE}")
    df = pd.read_excel(INPUT_FILE)

    print(f"\n========== 开始测试: {display_name} ({api_model_name}) ==========")

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

            print(f"[{display_name}] 进度: {finished}/{total}")

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
    sheet_name = display_name[:31]  # Excel Sheet 名字最长31个字符

    # 保存单个模型的结果
    out_df.to_excel(output_file, sheet_name=sheet_name, index=False)

    print(f"{display_name} 完成")
    print(f"✅ 结果已保存至: {output_file}")


if __name__ == "__main__":
    main()