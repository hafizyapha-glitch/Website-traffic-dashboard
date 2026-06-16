import os
import csv
import requests

def fetch_facebook_ads_data():
    # 1. ดึงข้อมูลจาก GitHub Secrets
    access_token = os.environ.get('FB_ACCESS_TOKEN')
    ad_account_id = os.environ.get('FB_OBJECT_ID')
    
    if not access_token or not ad_account_id:
        print("❌ ข้อผิดพลาด: ไม่พบ Access Token หรือ Ad Account ID กรุณาเช็ค GitHub Secrets")
        return False
        
    # 2. ตรวจสอบและเติม 'act_' หน้า Ad Account ID (ถ้ายังไม่มี)
    # Facebook Ads API บังคับให้ใส่คำว่า act_ นำหน้าไอดีบัญชีโฆษณาเสมอ
    if not ad_account_id.startswith('act_'):
        ad_account_id = f"act_{ad_account_id}"
        
    # 3. ตั้งค่า Endpoint สำหรับดึงสถิติ Facebook Ads (ระดับ Ad Set)
    url = f"https://graph.facebook.com/v18.0/{ad_account_id}/insights"
    
    params = {
        'level': 'adset', # ดึงข้อมูลลึกระดับชุดโฆษณา
        'breakdowns': 'age,gender', # แยกข้อมูลตามอายุและเพศ
        'fields': 'campaign_name,adset_name,spend,reach,impressions,actions,action_values',
        'date_preset': 'last_30d', # ดึงข้อมูลย้อนหลัง 30 วัน (เปลี่ยนเป็น 'maximum' หรือ 'this_month' ได้)
        'access_token': access_token
    }
    
    try:
        print(f"กำลังดึงข้อมูลโฆษณาจากบัญชี {ad_account_id}...")
        response = requests.get(url, params=params)
        response.raise_for_status() 
        data = response.json().get('data', [])
        
        if not data:
            print("⚠️ ดึงข้อมูลสำเร็จ แต่ไม่พบข้อมูลสถิติในช่วงเวลานี้")
        
        # 4. กำหนดหัวตารางให้ตรงกับที่ไฟล์ HTML ของคุณต้องการเป๊ะๆ
        headers = [
            "Campaign name", "Ad set name", "Age", "Gender", 
            "Amount spent (THB)", "Reach", "Impressions", 
            "Link clicks", "Landing page views", 
            "Website adds to cart", "Website purchases", 
            "Website purchases conversion value"
        ]
        
        os.makedirs('data', exist_ok=True)
        output_path = 'data/traffic-data.csv'
        
        # 5. เขียนข้อมูลลงไฟล์ CSV
        with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(headers)
            
            for row in data:
                # แยกย่อยข้อมูล Actions (พฤติกรรมต่างๆ)
                actions = row.get('actions', [])
                link_clicks = 0
                lp_views = 0
                adds_to_cart = 0
                purchases = 0
                
                for action in actions:
                    action_type = action.get('action_type')
                    value = int(action.get('value', 0))
                    if action_type == 'link_click':
                        link_clicks += value
                    elif action_type == 'landing_page_view':
                        lp_views += value
                    elif action_type == 'add_to_cart':
                        adds_to_cart += value
                    elif action_type == 'purchase':
                        purchases += value
                        
                # แยกย่อยข้อมูล Action Values (มูลค่าการสั่งซื้อ)
                action_values = row.get('action_values', [])
                purchase_value = 0.0
                for val in action_values:
                    if val.get('action_type') == 'purchase':
                        purchase_value += float(val.get('value', 0))
                        
                # บันทึกลงตาราง CSV
                writer.writerow([
                    row.get('campaign_name', 'Unknown'),
                    row.get('adset_name', 'Unknown'),
                    row.get('age', 'Unknown'),
                    row.get('gender', 'Unknown'),
                    row.get('spend', 0),
                    row.get('reach', 0),
                    row.get('impressions', 0),
                    link_clicks,
                    lp_views,
                    adds_to_cart,
                    purchases,
                    purchase_value
                ])
                
        print(f"✅ ดึงข้อมูลสำเร็จ! บันทึกไฟล์ไว้ที่: {output_path}")
        return True
        
    except requests.exceptions.RequestException as e:
        print(f"❌ เกิดข้อผิดพลาดจาก API: {e}")
        # ปริ้นท์รายละเอียด Error จาก Facebook เพื่อให้แก้ปัญหาได้ง่ายขึ้น
        if hasattr(e, 'response') and e.response is not None:
            print("รายละเอียด Error:", e.response.json())
        return False

if __name__ == "__main__":
    fetch_facebook_ads_data()
