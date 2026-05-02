import requests
import json
import os
from datetime import datetime

# 绝杀：直接请求知乎官方底层接口，跳过所有第三方中间商
API_URL = "https://www.zhihu.com/api/v3/feed/topstory/hot-lists/total?limit=50&desktop=true"

def fetch_and_save():
    print(f"开始抓取数据: {API_URL}")
    
    # 伪装成正常的电脑浏览器
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
        "Accept": "*/*",
        "Host": "www.zhihu.com"
    }
    
    try:
        response = requests.get(API_URL, headers=headers, timeout=15)
        print(f"服务器状态码: {response.status_code}")
        
        if response.status_code == 200:
            data = response.json()
            
            os.makedirs("data", exist_ok=True)
            today_str = datetime.now().strftime("%Y-%m-%d")
            filename = f"data/zhihu_{today_str}.json"
            
            with open(filename, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
                
            print(f"✅ 抓取成功！数据已保存至: {filename}")
        else:
            print(f"❌ 抓取失败！HTTP 状态码: {response.status_code}")
            print(f"返回内容: {response.text[:200]}")
            
    except Exception as e:
        print(f"❌ 运行发生严重错误: {e}")

if __name__ == "__main__":
    fetch_and_save()
