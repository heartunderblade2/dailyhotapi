import requests
import json
import os
from datetime import datetime

# 使用稳定且免费的韩小韩 API 接口
API_URL = "https://api.vvhan.com/api/hotlist/zhihu"

def fetch_and_save():
    print(f"开始抓取数据: {API_URL}")
    
    # 伪装成浏览器，防止被 API 拦截
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
    
    try:
        # 发送请求
        response = requests.get(API_URL, headers=headers, timeout=15)
        
        if response.status_code == 200:
            data = response.json()
            
            # 确保 data 文件夹存在
            os.makedirs("data", exist_ok=True)
            
            # 生成带日期的文件名
            today_str = datetime.now().strftime("%Y-%m-%d")
            filename = f"data/zhihu_{today_str}.json"
            
            # 保存为 JSON 文件
            with open(filename, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
                
            print(f"✅ 抓取成功！数据已保存至: {filename}")
        else:
            print(f"❌ 抓取失败！HTTP 状态码: {response.status_code}")
            
    except Exception as e:
        print(f"❌ 运行发生严重错误: {e}")

if __name__ == "__main__":
    fetch_and_save()
