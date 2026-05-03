import requests
import json
import os
from datetime import datetime, timezone, timedelta

# 你的专属 Vercel 基础 API 地址
BASE_URL = "https://my-dailyhotapi-vercel.vercel.app/"

# 把你想抓取的所有平台代号放在这个列表里
TARGETS = [
    "bilibili",
    "toutiao",
    "thepaper",
    "douyin",
    "douban-group",
    "ithome",
    "qq-news",
    "sina",
    "sina-news",
    "netease-news"
]

def fetch_and_save():
    # 划重点：在循环外统一获取一次时间，保证这批次的所有文件时间戳一模一样
    tz_beijing = timezone(timedelta(hours=8))
    now_str = datetime.now(tz_beijing).strftime("%Y-%m-%d_%H-%M")
    
    print(f"🚀 开始批量分类抓取数据，当前批次时间: {now_str}\n")
    print("-" * 40)
    
    # 开始循环遍历每一个平台
    for target in TARGETS:
        api_url = f"{BASE_URL}{target}"
        print(f"⏳ 正在抓取 [{target}] ...")
        
        try:
            # 发送网络请求
            response = requests.get(api_url, timeout=15)
            
            # 状态码 200 代表完全成功
            if response.status_code == 200:
                data = response.json()
                
                # 【新增逻辑】：为每个平台动态创建专属文件夹
                # 路径例如: data/bilibili, data/toutiao
                target_dir = f"data/{target}"
                os.makedirs(target_dir, exist_ok=True)
                
                # 拼装最终的文件保存路径
                # 结果例如: data/bilibili/bilibili_2026-05-03_14-30.json
                filename = f"{target_dir}/{target}_{now_str}.json"
                
                # 把抓下来的数据存进专属文件夹的 JSON 文件里
                with open(filename, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                    
                print(f"✅ 成功! 保存至 -> {filename}\n")
                
            else:
                print(f"❌ 失败! [{target}] HTTP 状态码: {response.status_code}\n")
                
        except Exception as e:
            # 即使某一个抓取崩溃，也会被这里拦截，保证下一个 target 继续执行
            print(f"❌ 崩溃! [{target}] 发生异常: {e}\n")
            
    print("-" * 40)
    print("🎉 所有平台的抓取任务执行完毕！")

if __name__ == "__main__":
    fetch_and_save()
