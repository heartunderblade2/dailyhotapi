import requests
import json
import os
from datetime import datetime, timezone, timedelta

# 你的专属 Vercel API 地址（请确保没有拼写错误）
API_URL = "https://my-dailyhotapi-vercel.vercel.app/bilibili"

def fetch_and_save():
    print(f"开始抓取数据: {API_URL}")
    
    try:
        # 发送最纯粹的网络请求
        response = requests.get(API_URL, timeout=15)
        print(f"服务器状态码: {response.status_code}")
        
        # 状态码 200 代表完全成功
        if response.status_code == 200:
            data = response.json()
            
            # 确保存放数据的 data 文件夹存在
            os.makedirs("data", exist_ok=True)
            
            # 划重点：强制获取北京时间 (UTC+8)
            tz_beijing = timezone(timedelta(hours=8))
            # 格式化为 年-月-日_时-分 (例如: 2026-05-03_14-30)
            now_str = datetime.now(tz_beijing).strftime("%Y-%m-%d_%H-%M")
            
            filename = f"data/bilibili_{now_str}.json"
            
            # 把抓下来的数据妥善存进 JSON 文件里
            with open(filename, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
                
            print(f"✅ 抓取成功！数据已完美保存至: {filename}")
            
        else:
            print(f"❌ 抓取失败！HTTP 状态码: {response.status_code}")
            print(f"服务器返回的错误提示: {response.text[:200]}")
            
    except Exception as e:
        print(f"❌ 代码执行中途崩溃，原因: {e}")

if __name__ == "__main__":
    fetch_and_save()
