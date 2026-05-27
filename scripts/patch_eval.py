import os
import sys
import time
import requests
import pandas as pd
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

# =========================
# 1. 基础配置
# =========================

OUTPUT_DIR = Path("outputs")
OUTPUT_DIR.mkdir(exist_ok=True)

API_BASE_URL = "https://api.moark.com/v1"
API_KEY = os.getenv("API_KEY")

if not API_KEY:
    raise ValueError("未设置 API_KEY 环境变量，请检查配置！")

# 这里保留映射关系，用于从 Sheet 名或文件名匹配对应的 API 模型名
MODELS = {
    "Deepseek-R1": "deepseek-r1",
    "GLM-5": "glm-5",
    "Kimi-K2.5": "Kimi-K2.5",
    "Qwen3-235B-A22B": "qwen3-235b-a22b",
    "GPT-4o-mini": "openai/gpt-4o-mini",
    # "Gemini-2.5-Flash-Lite": "google/gemini-2.5-flash-lite",
}

MAX_WORKERS = 2  # 并发数

# =========================
# 2. API 调用函数 (保持不变)
# =========================
def call_model(model_name: str, prompt: str, retries: int = 2) -> str:
    if pd.isna(prompt) or not str(prompt).strip():
        return ""

    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json",
    }

    payload = {
        "model": model_name,
        "messages": [{"role": "user", "content": str(prompt)}],
        "temperature": 0.1,
        "max_tokens": 4096,
    }

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
                print(f"[{model_name}] HTTP {response.status_code}")
                if response.status_code == 400:
                    return f"[HTTP 400] {response.text}"
                time.sleep(2)
        except Exception as e:
            print(f"[异常] {model_name}: {e}")
            time.sleep(2)

    return "[请求失败] 超过最大重试次数"

# =========================
# 3. 补漏单行处理逻辑
# =========================
def process_patch_row(idx, row, api_model_name):
    """
    检查该行是否需要补齐英文输出。
    只重跑确实需要跑的数据，其他原样返回。
    """
    updated_row = row.copy()
    
    en_prompt = updated_row.get("英文prompt", "")
    en_output = updated_row.get("对应的输出(英文)", "")

    # 判断条件：有 Prompt，但 Output 是空的、NaN 或者包含之前的失败标记
    has_prompt = pd.notna(en_prompt) and str(en_prompt).strip() != ""
    is_output_missing = (
        pd.isna(en_output) 
        or str(en_output).strip() == "" 
        or "失败" in str(en_output) 
        or "异常" in str(en_output)
    )

    if has_prompt and is_output_missing:
        # 只针对缺失/失败的英文进行补调
        new_en_output = call_model(api_model_name, str(en_prompt))
        updated_row["对应的输出(英文)"] = new_en_output
        time.sleep(0.5)
        return idx, updated_row, True  # True 表示执行了更新

    return idx, updated_row, False # False 表示跳过（已存在）

# =========================
# 4. 主逻辑
# =========================
def main():
    # 动态获取输入文件
    if len(sys.argv) < 2:
        print("用法: python patch_eval.py <输入文件.xlsx>")
        sys.exit(1)

    input_file = Path(sys.argv[1])
    if not input_file.exists():
        print(f"❌ 找不到输入文件: {input_file}")
        sys.exit(1)

    output_file = OUTPUT_DIR / f"{input_file.stem}_fixed.xlsx"
    print(f"正在加载 Excel 数据: {input_file} ...")
    
    # 读取所有 Sheet
    df_dict = pd.read_excel(input_file, sheet_name=None)
    
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        
        for sheet_name, df in df_dict.items():
            print(f"\n========== 正在检查 Sheet: {sheet_name} ==========")
            
            # 从 Sheet 名映射 API 模型名 (兼容原来的 MODELS 字典)
            # 如果你的 sheet 名字就是 'Deepseek-R1' 这种，就能直接取到
            api_model_name = MODELS.get(sheet_name)
            
            # 如果匹配不到（比如 sheet 名字变了），尝试模糊匹配
            if not api_model_name:
                for key, val in MODELS.items():
                    if key.lower() in sheet_name.lower():
                        api_model_name = val
                        break
            
            # 如果依然找不到，要求检查配置
            if not api_model_name:
                print(f"⚠️ 无法为 Sheet '{sheet_name}' 匹配到对应的 api_model_name，跳过此 Sheet。")
                df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
                continue
                
            print(f"使用的模型接口: {api_model_name}")

            results = {}
            updated_count = 0

            with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
                futures = []
                for idx, row in df.iterrows():
                    future = executor.submit(process_patch_row, idx, row, api_model_name)
                    futures.append(future)

                total = len(futures)
                finished = 0

                for future in as_completed(futures):
                    finished += 1
                    try:
                        idx, result_row, was_updated = future.result()
                        results[idx] = result_row
                        if was_updated:
                            updated_count += 1
                    except Exception as e:
                        print(f"[任务异常]: {e}")
                    
                    if finished % 50 == 0 or finished == total:
                        print(f"[{sheet_name}] 进度: {finished}/{total}")

            print(f"[{sheet_name}] 完成。本次共修复补充了 {updated_count} 条数据。")

            # 按原始顺序重建 DataFrame
            rows = [results[idx] for idx in df.index if idx in results]
            out_df = pd.DataFrame(rows)
            
            out_df.to_excel(writer, sheet_name=sheet_name[:31], index=False)

    print(f"\n✅ 全部补漏完成！结果已保存至: {output_file}")

if __name__ == "__main__":
    main()
