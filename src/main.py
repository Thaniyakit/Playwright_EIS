import os
import time
import pandas as pd
from playwright.sync_api import sync_playwright
from dotenv import load_dotenv
from datetime import datetime, timedelta
from bs4 import BeautifulSoup

# โหลดค่าจากไฟล์ .env
load_dotenv()

OUTPUT_DIR = os.getenv("OUTPUT_DIR", "./data/raw_data") + datetime.now().strftime("%m-%Y")
CLEANED_DIR = os.getenv("CLEANED_DIR", "./data/cleaned_data") + datetime.now().strftime("%m-%Y")
#output_dir + month + year = datetime.now().strftime("%Y-%m")
os.makedirs(OUTPUT_DIR , exist_ok=True)
os.makedirs(CLEANED_DIR , exist_ok=True)


TARGET = [ {"name": "ร้านNewPB", "id": "287"},
		   {"name": "ร้านNewSA", "id": "263"}, 
		   {"name": "ร้านNewSH", "id": "285"}, 
		   {"name": "ร้านRenovate", "id": "235"} ]

def calculate_last_month_date_range():
	# 1. ดึงวันที่ปัจจุบัน
	today = datetime.now()

	# 2. คำนวณหา "วันแรกของเดือนก่อนหน้า"
	# โดยเอาวันแรกของเดือนปัจจุบัน ลบออก 1 วัน จะได้วันสุดท้ายของเดือนก่อนเสมอ
	first_day_of_this_month = today.replace(day=1)
	last_day_of_prev_month = first_day_of_this_month - timedelta(days=1)

	# จากนั้นกำหนดวันเป็นวันที่ 1 ของเดือนนั้น
	first_day_of_prev_month = last_day_of_prev_month.replace(day=1)

	# 3. จัด Format วันที่ให้อยู่ในรูปแบบ D/M/YYYY (เช่น 1/7/2026 หรือ 31/7/2026)
	# หมายเหตุ: หากหน้าเว็บรองรับปี พ.ศ. (บวก 543) ให้บวกเพิ่มที่ .year
	is_buddhist_year = True  # เปลี่ยนเป็น False หากหน้าเว็บใช้ปี ค.ศ.
	year_offset = 543 if is_buddhist_year else 0

	start_year = first_day_of_prev_month.year + year_offset
	end_year = last_day_of_prev_month.year + year_offset

	date_from_str = f"{first_day_of_prev_month.day}/{first_day_of_prev_month.month}/{start_year}"
	date_to_str = f"{last_day_of_prev_month.day}/{last_day_of_prev_month.month}/{end_year}"

	return date_from_str, date_to_str

def read_xml_xls(filepath):
    """ฟังก์ชันพิเศษสำหรับแกะไฟล์ .xls ที่เนื้อหาข้างในเป็น Excel XML (SpreadsheetML)"""
    with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
        soup = BeautifulSoup(f.read(), "xml")

    # ดึงทุก Row ในตาราง XML
    rows = soup.find_all("Row")
    data = []

    for row in rows:
        # ดึงข้อความในแต่ละ Cell ของแถวนั้นๆ
        cells = [cell.text.strip() for cell in row.find_all("Cell")]
        if cells:
            data.append(cells)

    if not data:
        return None

    # แถวแรกคือ Header (ชื่อคอลัมน์)
    headers = data[0]
    # แถวที่เหลือคือ Data
    row_data = data[1:]

    # สร้าง DataFrame
    df = pd.DataFrame(row_data, columns=headers)
    return df

def clean_and_transform(df: pd.DataFrame, filename: str, outputfilename: str) -> pd.DataFrame:
    """ฟังก์ชัน Data Cleaning & Column Dropping ด้วย Pandas"""
    print("[*] Cleaning and formatting data...")

    filepath = os.path.join(OUTPUT_DIR, f"{filename}_{datetime.now().strftime('%m-%Y')}.xls")
    print(f"[*] Reading data from: {filepath}")
    df = None

    # 1. ลองอ่านแบบ XML / HTML Table ก่อน (เพราะระบบองค์กรส่วนใหญ่เป็น XML/HTML ปลอมนามสกุลเป็น .xls)
    try:
        df = read_xml_xls(filepath)
    except Exception as e:
        print(f"[!] แกะ XML ไม่สำเร็จ: {e}")

	# ตัด Column ที่ไม่ใช้ออก (เลือกเก็บเฉพาะ Column ที่ต้องการ)
    cols_to_keep = [
		"รหัสสาขา",
		"ชื่อสาขา",
		"สถานะ",
		"บ้านเลขที่",
		"อาคาร/หมู่บ้าน",
		"ซอย/ถนน",
		"รหัสทำเล",
		"-",
		"จังหวัด",
		"อำเภอ/เขต",
		"ตำบล/แขวง",
		"รหัสไปรษณีย์",
		"พื้นที่",
		"ประเภทร้าน",
		"ระบุวันที่เข้าเคลียร์ไซต์จริง",
		"วันเคลียร์ไซต์จากวงบริหารแผนเปิดร้าน",
		"ป้ายแสดงระยะทาง",
		"ป้ายแสดงระยะทาง โครงไม้/โครงเหล็ก",
		"จำนวนป้ายแสดงระยะทาง",
		"ป้ายที่ 1",
		"ป้ายที่ 2",
		"ป้ายที่ 3",
		"ป้ายที่ 4",
		"ป้ายที่ 5",
		"ป้ายที่ 6",
		"ป้ายที่ 7",
		"ป้ายที่ 8",
		"ป้ายที่ 9",
		"ป้ายที่ 10",
		"แนบไฟล์ PDF. (แผนที่พร้อมระบุตำแหน่งติดตั้งและระยะทาง)",
		"ป้ายแสดงข้อความ",
		"ป้ายแสดงข้อความ โครงไม้/โครงเหล็ก",
		"จำนวนป้ายแสดงข้อความ",
		"ป้ายที่ 1 ระบุข้อความ...",
		"ป้ายที่ 2 ระบุข้อความ...",
		"ป้ายที่ 3 ระบุข้อความ...",
		"ป้ายที่ 4 ระบุข้อความ...",
		"ป้ายที่ 5 ระบุข้อความ...",
		"ป้ายที่ 6 ระบุข้อความ...",
		"ป้ายที่ 7 ระบุข้อความ...",
		"ป้ายที่ 8 ระบุข้อความ...",
		"ป้ายที่ 9 ระบุข้อความ...",
		"ป้ายที่ 10 ระบุข้อความ...",
		"แนบไฟล์ PDF. (แผนที่พร้อมระบุตำแหน่งติดตั้งและระยะทาง)",
	]
    if "รหัสทำเล" not in df.columns:
        cols_to_keep.remove("รหัสทำเล")
        print("[!] Column 'รหัสทำเล' ไม่มีข้อมูลทั้งหมด จะไม่ถูกเก็บไว้ในไฟล์")
    if "-" not in df.columns:
        cols_to_keep.remove("-")
        print("[!] Column '-' ไม่มีข้อมูลทั้งหมด จะไม่ถูกเก็บไว้ในไฟล์")
    df_cleaned = df[cols_to_keep].copy()

    # Data formatting
    # df_cleaned["Product_Name"] = df_cleaned["Product_Name"].astype(str).str.strip()
    # df_cleaned["Price"] = (
    #     df_cleaned["Price"]
    #     .astype(str)
    #     .str.replace("$", "", regex=False)
    #     .astype(float)
    # )

    df_cleaned.to_excel(
		CLEANED_DIR + "/" + outputfilename + ".xlsx", index=False, engine="openpyxl"
	)

    return df_cleaned

def scrape_data():
	with sync_playwright() as p:
	
			browser = p.chromium.launch(
				headless=True,
				# แนะนำให้เพิ่ม '--start-maximized' เพื่อเปิดหน้าจอใหญ่เต็มจอ
				args=["--start-maximized"]
			)
			
			# สร้าง context แบบกำหนดขนาดหน้าจอให้เต็มขอบ
			context = browser.new_context(no_viewport=True)
			page = context.new_page()
	
			print("กำลังเปิดเว็บ...")
			page.goto("https://eis4ce.cpall.co.th/bench", timeout=60000) # เปลี่ยน URL เป็นเว็บไซต์ที่ต้องการดึงข้อมูล
			page.locator('#Username').fill(os.getenv("SCRAPER_USERNAME"))
			print(os.getenv("SCRAPER_USERNAME"))
			page.locator('#Password').fill(os.getenv("SCRAPER_PASSWORD"))
			print(os.getenv("SCRAPER_PASSWORD"))
			page.get_by_role("button", name="เข้าสู่ระบบ").click()
			print("เข้าสู่ระบบเรียบร้อยแล้ว")
	
			page.wait_for_load_state("networkidle", timeout=60000)  # รอให้หน้าเว็บโหลดเสร็จ
			print("โหลดหน้าเว็บแล้ว")
			page.locator("a").filter(has_text="ค้นหา").first.click()
			page.locator("#searchExtraFormLink").click()
			for i in range(len(TARGET)):
				page.locator("#ExtraFormDropdown").select_option(TARGET[i]["id"])
				LastMonth = time.localtime().tm_mon - 1
				start_date, end_date = calculate_last_month_date_range()
				print(f"กำลังกรอกวันที่: {start_date} ถึง {end_date}")
				page.locator("#CreatedDateFrom").fill(start_date)
				page.keyboard.press("Escape")
				page.locator("#CreatedDateTo").fill(end_date)
				time.sleep(1)  # รอให้หน้าเว็บโหลดข้อมูล
				page.keyboard.press("Escape")
				page.locator("#searchExtraFormButton").click()
				time.sleep(5)  # รอให้หน้าเว็บโหลดข้อมูล
				page.get_by_role("img", name="ส่งออกข้อมูล").click()
				print("กดส่งออกข้อมูลเรียบร้อยแล้ว")
				# ชี้ไปที่ ID โดยตรง แล้วสั่งคลิก
				page.locator("#excelFormDetail").click()
				print("เลือกฟอร์มพิเศษเรียบร้อยแล้ว")
				page.locator("#excelFormDivselect").select_option(TARGET[i]["id"])
	
				with page.expect_download() as download_info:
					# สั่งกดปุ่มดาวน์โหลดภายในบล็อกนี้
					page.locator("div[id='excelFormDiv'] div div input[value='ส่งออก']").click(force=True)
	
				# 3. ดึง Object ไฟล์ที่กำลังดาวน์โหลดมา
				download = download_info.value
	
				save_path = os.path.join(OUTPUT_DIR, f"{TARGET[i]['name']}_{datetime.now().strftime('%m-%Y')}.xls")
	
				# 4. บันทึกไฟล์ลงในโฟลเดอร์ ./data/
				download.save_as(save_path)
	
				print(f"[✔] ดาวน์โหลดไฟล์สำเร็จ! บันทึกไว้ที่: {save_path}")
	
				page.locator("#closeExcelFormDiv").click()
				page.locator("#closeExportDiv").click()
				time.sleep(2)  # รอให้หน้าเว็บโหลดข้อมูล
			# browser.close()
			  # รอให้หน้าเว็บโหลดข้อมูล
def main() :
    scrape_data()
    for i in range(len(TARGET)):
        filename = TARGET[i]["name"]
        outputfilename = filename + "_" + datetime.now().strftime("%m-%Y")
        clean_and_transform(None, filename, outputfilename)


if __name__ == "__main__":
    main()