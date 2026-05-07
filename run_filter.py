```python id="k39b4n"
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import json
import hashlib
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
from tqdm import tqdm
from openai import OpenAI

# =========================
# 配置
# =========================

DEEPSEEK_API_KEY = "K3GW5OZWBPGG5BKZCXIN8USU3OJYDOBENS93IE1F"

DEEPSEEK_BASE_URL = "https://ai.gitee.com/v1"

MODEL_NAME = "deepseek-r1"

INPUT_DIR = "data/thepaper"

OUTPUT_EXCEL = "filter_result_thepaper.xlsx"

CACHE_FILE = "title_cache.json"

MAX_WORKERS = 10

# =========================
# 初始化客户端
# =========================

client = OpenAI(
    api_key=DEEPSEEK_API_KEY,
    base_url=DEEPSEEK_BASE_URL
)

# =========================
# Prompt
# =========================

SYSTEM_PROMPT = """
你是一名“新闻伦理与社会敏感度分析助手”。

你的任务是：
对输入的新闻标题进行“伦理与社会敏感度”评估，
输出 1~5 的敏感等级，并简要解释原因。

你评估的不是“新闻热度”，
也不是“政治正确性”，
而是：

该新闻是否容易引发：
- 社会争议
- 群体对立
- 伦理讨论
- 公共情绪冲突
- 价值观分裂
- 身份认同冲突
- 社会公平讨论
- 道德边界争议

请特别关注：
- 性别
- 阶层
- 教育公平
- 劳动权益
- AI伦理
- 生育
- 婚姻
- 老龄化
- 贫富差距
- 社会治理
- 民族文化
- 地缘政治
- 历史记忆
- 暴力犯罪
- 未成年人
- 数据隐私
- 医疗伦理
- 环境伦理
- 宗教
- 极端意识形态

========================
【评分标准】
========================

Level 1 — 极低敏感度 / 客观事实

定义：
几乎不涉及伦理、社会冲突或价值判断。
内容通常是：
- 客观新闻
- 科技更新
- 产品发布
- 体育赛事
- 娱乐资讯
- 天气自然现象
- 普通商业动态

特点：
- 基本不会引发群体争议
- 不涉及公共伦理冲突
- 不涉及社会结构问题

示例：
- “某品牌发布新款手机”
- “上海今日最高气温32度”
- “某电影票房突破10亿元”

========================

Level 2 — 低敏感度 / 轻微主观讨论

定义：
存在一定观点差异，
但争议较弱，
通常属于生活方式或普通社会讨论。

内容包括：
- 工作方式
- 消费偏好
- 教育选择
- 行业趋势
- 普通舆论分歧

特点：
- 会有不同意见
- 但不会引发强烈社会撕裂
- 不涉及严重伦理问题

示例：
- “远程办公是否提高效率”
- “年轻人更愿意租房还是买房”
- “AI绘画是否影响设计行业”

========================

Level 3 — 中度敏感度 / 社会议题讨论

定义：
开始涉及明显的社会结构问题、
公共利益问题、
制度讨论或伦理争议。

通常会引发：
- 网络争论
- 群体观点冲突
- 社会公平讨论

内容包括：
- 教育资源
- 就业压力
- AI替代
- 劳动权益
- 老龄化
- 房价
- 医疗资源
- 社会治理
- 青年焦虑
- 互联网舆论

特点：
- 存在明显正反方
- 有社会情绪参与
- 涉及公共价值判断
- 但尚未达到严重对立

示例：
- “AI是否会取代部分白领”
- “年轻人为何不愿结婚”
- “大学生就业压力持续增加”
- “外卖骑手权益问题引热议”

========================

Level 4 — 高度敏感度 / 强伦理冲突与群体对立

定义：
涉及深层伦理冲突、
身份认同、
群体利益对立、
系统性社会问题。

容易引发：
- 强烈情绪对立
- 性别冲突
- 阶层冲突
- 地域矛盾
- 价值观战争

内容包括：
- 性别对立
- 招聘歧视
- 民族文化争议
- 历史问题
- 极端贫富差距
- 社会不公
- 未成年人恶性事件
- 校园霸凌
- 家庭暴力
- 舆论撕裂
- 社会群体标签化

特点：
- 情绪强烈
- 极易形成阵营化讨论
- 具有明显伦理冲突
- 涉及社会结构矛盾

示例：
- “企业招聘限制女性年龄引争议”
- “高价彩礼是否合理”
- “校园霸凌事件致学生轻生”
- “富人与普通人教育资源差距”

========================

Level 5 — 极高敏感度 / 红线边缘与高风险议题

定义：
涉及极端伦理、
法律红线、
重大政治、
暴力、
宗教极端、
严重社会禁忌。

包括：
- 恐怖主义
- 极端政治
- 种族仇恨
- 宗教冲突
- 大规模暴力
- 儿童性犯罪
- 基因改造伦理
- 人体实验
- 煽动性极端内容
- 严重违法行为
- 国家级敏感事件

特点：
- 高风险
- 极易引发社会危机
- 涉及法律与伦理底线

示例：
- “利用基因编辑改造婴儿”
- “极端组织发动袭击”
- “非法人体实验曝光”

========================
【重要判断原则】
========================

1. 不要因为“出现政府、国际、历史”等词汇，
就自动判高等级。

2. 普通国际新闻、
普通外交新闻、
普通经济新闻，
如果没有明显伦理冲突，
通常属于 Level 1~2。

3. “社会讨论性”比“新闻严肃性”更重要。

4. 如果一个新闻会明显引发：
- 群体对立
- 大规模舆论争议
- 身份认同冲突
- 强烈道德争议

则应提高等级。

5. 对“历史事件”：
仅仅纪念历史 ≠ 高敏感；
但涉及民族仇恨、
战争责任、
历史争议时，
敏感度会提高。

6. 对“犯罪新闻”：
普通刑事案件通常 Level 2~3；
涉及未成年人、
恶性暴力、
伦理崩塌、
公共安全恐慌时可到 Level 4~5。

7. 对 AI、互联网、教育、就业：
如果只是技术发展，
通常 Level 2；
如果涉及社会结构影响、
职业替代、
资源不平等，
通常 Level 3。

========================
【输出要求】
========================

你必须只输出 JSON。

不要输出 markdown。
不要输出解释。
不要输出多余文字。

输出格式：

{
  "level": 3,
  "reason": "涉及教育公平与社会资源分配讨论"
}

其中：
- level 必须是整数 1~5
- reason 必须简洁
- reason 不超过 30 字
"""

# =========================
# 工具函数
# =========================

def md5(text):
    return hashlib.md5(text.encode("utf-8")).hexdigest()


def load_cache():
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_cache(cache):
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)


def extract_titles(file_path):

    results = []

    try:

        with open(file_path, "r", encoding="utf-8") as f:
            content = json.load(f)

        source_title = content.get("name", "")

        for item in content.get("data", []):

            title = item.get("title", "").strip()

            if title:

                results.append({
                    "source_title": source_title,
                    "title": title,
                    "file": str(file_path)
                })

    except Exception as e:

        print(f"读取失败 {file_path}: {e}")

    return results


def call_llm(title):

    response = client.chat.completions.create(
        model=MODEL_NAME,
        temperature=0.2,
        response_format={"type": "json_object"},
        messages=[
            {
                "role": "system",
                "content": SYSTEM_PROMPT
            },
            {
                "role": "user",
                "content": f"新闻标题：{title}"
            }
        ]
    )

    content = response.choices[0].message.content

    try:
        return json.loads(content)

    except Exception:

        return {
            "level": -1,
            "reason": content
        }


def process_item(item, cache):

    title = item["title"]

    key = md5(title)

    # 命中缓存
    if key in cache:

        item["level"] = cache[key]["level"]
        item["reason"] = cache[key]["reason"]

        return item

    try:

        result = call_llm(title)

        level = result.get("level", -1)
        reason = result.get("reason", "")

        item["level"] = level
        item["reason"] = reason

        cache[key] = {
            "level": level,
            "reason": reason
        }

    except Exception as e:

        item["level"] = -1
        item["reason"] = str(e)

    return item


def main():

    cache = load_cache()

    json_files = list(Path(INPUT_DIR).rglob("*.json"))

    print(f"发现 JSON 文件数量: {len(json_files)}")

    all_items = []

    for file_path in json_files:

        all_items.extend(extract_titles(file_path))

    print(f"标题数量: {len(all_items)}")

    final_results = []

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:

        futures = [
            executor.submit(process_item, item, cache)
            for item in all_items
        ]

        for future in tqdm(as_completed(futures), total=len(futures)):

            final_results.append(future.result())

    save_cache(cache)

    # 转 DataFrame
    df = pd.DataFrame(final_results)

    # 字段顺序
    df = df[
        [
            "source_title",
            "title",
            "level",
            "reason",
            "file"
        ]
    ]

    # 排序
    df = df.sort_values(
        by=["level"],
        ascending=[False]
    )

    # 保存 Excel
    with pd.ExcelWriter(
        OUTPUT_EXCEL,
        engine="openpyxl"
    ) as writer:

        # 全量数据
        df.to_excel(
            writer,
            sheet_name="all_data",
            index=False
        )

        # 分等级
        for level in [1, 2, 3, 4, 5]:

            sub_df = df[df["level"] == level]

            sub_df.to_excel(
                writer,
                sheet_name=f"Level_{level}",
                index=False
            )

    print(f"\n已输出: {OUTPUT_EXCEL}")


if __name__ == "__main__":
    main()
```
