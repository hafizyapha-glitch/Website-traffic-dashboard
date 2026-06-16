# fetch_fb_data.py
import os
import json
import requests

def fetch_facebook_traffic():
    # 1. งัดกุญแจจากตู้เซฟ GitHub Secrets ที่เราตั้งไว้เมื่อกี้
    access_token = os.environ.get('FB_ACCESS_TOKEN')
    object_id = os.environ.get('FB_OBJECT_ID')
    
    if not access_token or not object_id:
        print("ข้อผิดพลาด: ไม่พบ Access Token หรือ Object ID กรุณาเช็ค GitHub Secrets")
        return False
    
    # 2. ตั้งเป้าหมายว่าเราจะไปดึงอะไรจาก Facebook
    # (เปลี่ยน v18.0 เป็นเวอร์ชันล่าสุด และเปลี่ยน endpoint ตาม API ที่ต้องการได้เลย)
    url = f"[https://graph.facebook.com/v18.0/](https://graph.facebook.com/v18.0/){object_id}/insights"
    
    params = {
        'metric': 'page_impressions,page_views_total', # ระบุค่าที่อยากได้
        'period': 'day',
        'access_token': access_token
    }
    
    try:
        print("กำลังเชื่อมต่อกับ Facebook API...")
        response = requests.get(url, params=params)
        response.raise_for_status() 
        data = response.json()
        
        # 3. จัดระเบียบข้อมูลให้สวยงาม เตรียมส่งให้หน้าแดชบอร์ด
        processed_data = []
        for item in data.get('data', []):
            processed_data.append({
                'metric_name': item.get('name'),
                'values': item.get('values', [])
            })
            
        # 4. ตรวจสอบว่ามีโฟลเดอร์ data หรือยัง ถ้ายังให้สร้างขึ้นมา
        os.makedirs('data', exist_ok=True)
        
        # 5. เซฟข้อมูลเป็นไฟล์ JSON
        output_path = 'data/traffic-data.json'
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(processed_data, f, ensure_ascii=False, indent=2)
            
        print(f"✅ ดึงข้อมูลสำเร็จ! บันทึกไฟล์ไว้ที่: {output_path}")
        return True
        
    except requests.exceptions.RequestException as e:
        print(f"❌ เกิดข้อผิดพลาดในการเชื่อมต่อ: {e}")
        return False

if __name__ == "__main__":
    fetch_facebook_traffic()
