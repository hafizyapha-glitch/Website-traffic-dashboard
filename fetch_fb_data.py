import os
import csv
import requests
import json

def fetch_facebook_ads_data():
    # 1. ดึงข้อมูลตัวแปรลับ (Secrets) ที่เราตั้งไว้บน GitHub
    access_token = os.environ.get('FB_ACCESS_TOKEN')
    ad_account_id = os.environ.get('FB_OBJECT_ID')
    
    if not access_token or not ad_account_id:
        print("❌ ข้อผิดพลาด: ไม่พบ Access Token หรือ Ad Account ID กรุณาเช็คการตั้งค่าใน GitHub Secrets")
        return False
        
    # 2. ตรวจสอบรูปแบบ ID บัญชีโฆษณา (ต้องมีคำว่า 'act_' นำหน้าตามกฎของ Facebook API)
    if not ad_account_id.startswith('act_'):
        ad_account_id = f"act_{ad_account_id}"
        
    # 3. กำหนด URL ปลายทางสำหรับดึงข้อมูลระดับชุดโฆษณา (Ad Set)
    url = f"https://graph.facebook.com/v18.0/{ad_account_id}/insights"
    
    # 4. ตั้งค่าเงื่อนไขในการดึงข้อมูล
    params = {
        'level': 'adset', # ดึงข้อมูลลึกลงไประดับ Ad Set
        'breakdowns': 'age,gender', # ขอให้ Facebook แยกข้อมูลตามช่วงอายุและเพศมาให้ด้วย
        'fields': 'campaign_name,adset_name,spend,reach,impressions,actions,action_values', # ระบุคอลัมน์ที่ต้องการ
        'date_preset': 'last_30d', # ดึงข้อมูลสถิติย้อนหลัง 30 วัน
        'access_token': access_token
    }
    
    try:
        print(f"กำลังส่งคำขอดึงข้อมูลจากบัญชีโฆษณา: {ad_account_id} ...")
        response = requests.get(url, params=params)
        
        # 5. ตรวจสอบผลลัพธ์ว่า API ตอบกลับสำเร็จ (Status 200) หรือไม่
        if response.status_code != 200:
            print(f"❌ เกิดข้อผิดพลาดจากฝั่ง Facebook API (Status Code: {response.status_code})")
            # ปริ้นท์รายละเอียด Error แบบเจาะลึก เพื่อให้เราแก้ปัญหาได้ตรงจุด
            print("รายละเอียด Error:", json.dumps(response.json(), indent=2, ensure_ascii=False))
            return False
            
        # ดึงข้อมูลส่วนที่เป็น 'data' ออกมา
        data = response.json().get('data', [])
        
        if not data:
            print("⚠️ ดึงข้อมูลสำเร็จ! แต่... ไม่พบข้อมูลสถิติหรือยอดการใช้งานใดๆ ในช่วงเวลาที่เลือก (last_30d)")
            # โปรแกรมจะทำงานต่อเพื่อสร้างโครงตาราง CSV เปล่าๆ ป้องกันหน้าเว็บแสดงผลผิดพลาด
        
        # 6. กำหนดหัวคอลัมน์ของไฟล์ CSV (ต้องให้ตรงกับที่หน้าเว็บ Dashboard ดึงไปใช้)
        headers = [
            "Campaign name", "Ad set name", "Age", "Gender", 
            "Amount spent (THB)", "Reach", "Impressions", 
            "Link clicks", "Landing page views", 
            "Website adds to cart", "Website purchases", 
            "Website purchases conversion value"
        ]
        
        # 7. สร้างโฟลเดอร์สำหรับเก็บไฟล์ข้อมูล (หากในโปรเจกต์ยังไม่มีโฟลเดอร์ 'data')
        os.makedirs('data', exist_ok=True)
        output_path = 'data/traffic-data.csv'
        
        # 8. สร้างและเขียนข้อมูลลงไฟล์ CSV (ใช้ encoding='utf-8-sig' เพื่อให้ Excel อ่านภาษาไทยได้ไม่เพี้ยน)
        with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(headers) # เขียนหัวตารางลงไปก่อน 1 บรรทัด
            
            # วนลูปอ่านข้อมูลทีละแถว (ทีละ Ad Set/อายุ/เพศ) ที่ได้จาก Facebook
            for row in data:
                # 8.1 จัดการข้อมูลกลุ่ม 'Actions' (พฤติกรรม เช่น คลิก, หยิบใส่ตะกร้า)
                # Facebook ส่งมาเป็นก้อน Array เราต้องวนลูปหาประเภทที่ต้องการ
                actions = row.get('actions', [])
                link_clicks = 0
                lp_views = 0
                adds_to_cart = 0
                purchases = 0
                
                for action in actions:
                    action_type = action.get('action_type')
                    value = int(action.get('value', 0)) # แปลงค่าเป็นตัวเลข
                    if action_type == 'link_click':
                        link_clicks += value
                    elif action_type == 'landing_page_view':
                        lp_views += value
                    elif action_type == 'add_to_cart':
                        adds_to_cart += value
                    elif action_type == 'purchase':
                        purchases += value
                        
                # 8.2 จัดการข้อมูลกลุ่ม 'Action Values' (มูลค่าเป็นตัวเงิน เช่น ยอดขายรวม)
                action_values = row.get('action_values', [])
                purchase_value = 0.0
                for val in action_values:
                    if val.get('action_type') == 'purchase':
                        purchase_value += float(val.get('value', 0)) # แปลงเป็นตัวเลขทศนิยม
                        
                # 8.3 นำข้อมูลที่จัดระเบียบแล้ว เขียนลงไปในตาราง CSV ทีละบรรทัด
                writer.writerow([
                    row.get('campaign_name', 'ไม่ระบุชื่อแคมเปญ'),
                    row.get('adset_name', 'ไม่ระบุชื่อชุดโฆษณา'),
                    row.get('age', 'ไม่ระบุ'),
                    row.get('gender', 'ไม่ระบุ'),
                    row.get('spend', 0), # ค่าใช้จ่าย
                    row.get('reach', 0), # การเข้าถึง
                    row.get('impressions', 0), # การแสดงผล
                    link_clicks,
                    lp_views,
                    adds_to_cart,
                    purchases,
                    purchase_value
                ])
                
        print(f"✅ บันทึกข้อมูลเรียบร้อยแล้ว! เตรียมพบไฟล์ได้ที่: {output_path}")
        return True
        
    except requests.exceptions.RequestException as e:
        print(f"❌ ระบบเครือข่ายมีปัญหา ไม่สามารถเชื่อมต่อ Facebook ได้: {e}")
        return False
    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดของระบบระหว่างการประมวลผลข้อมูล: {e}")
        return False

# จุดเริ่มต้นการทำงานของโปรแกรม
if __name__ == "__main__":
    fetch_facebook_ads_data()
