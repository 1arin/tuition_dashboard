import asyncio
from playwright.async_api import async_playwright
import pandas as pd
from datetime import datetime

class MyTCASScraper:
    def __init__(self, base_url="https://course.mytcas.com"):
        self.base_url = base_url
        self.collected_data = []

    async def perform_search(self, page, keyword):
        """ค้นหาและเก็บข้อมูลรายการหลักสูตรจากคำค้น"""
        try:
            print(f"\n🔎 Searching for keyword: '{keyword}'")
            await page.goto(self.base_url, wait_until='networkidle')
            await asyncio.sleep(2)

            # หาช่องค้นหาแบบเจาะจง โดยลองหลาย selectors เผื่อหน้าเว็บเปลี่ยนแปลง
            search_input = None
            possible_selectors = [
                "input[placeholder='พิมพ์ชื่อมหาวิทยาลัย คณะ หรือหลักสูตร']",
                "input[type='search']",
                "input.search-input",
                "input[aria-label*='ค้นหา']",
            ]
            for sel in possible_selectors:
                try:
                    search_input = await page.query_selector(sel)
                    if search_input:
                        print(f"  ✔️ Found search input by selector: {sel}")
                        break
                except:
                    continue
            if not search_input:
                print("  ❌ ไม่พบช่องค้นหาในหน้าเว็บ")
                return []

            # ล้างค่าเดิมแล้วกรอกคำค้นหา
            await search_input.fill("")
            await asyncio.sleep(0.5)
            await search_input.fill(keyword)
            await search_input.press("Enter")
            await asyncio.sleep(3)

            # ดึงรายการผลลัพธ์ (ลองหลาย selector)
            program_list = []
            result_selectors = [
                "ul.t-programs > li",
                "ul.program-list > li",
                "div.search-results > ul > li",
                "[data-testid='program-item']",
            ]
            for sel in result_selectors:
                try:
                    items = await page.query_selector_all(sel)
                    if items and len(items) > 0:
                        print(f"  ✔️ พบ {len(items)} ผลลัพธ์โดยใช้ selector: {sel}")
                        program_list = items
                        break
                except:
                    continue
            if not program_list:
                print("  ❌ ไม่พบผลลัพธ์จากการค้นหา")
                return []

            programs = []
            for idx, item in enumerate(program_list):
                try:
                    text = await item.inner_text()
                    link_el = await item.query_selector("a")
                    if not link_el:
                        continue
                    href = await link_el.get_attribute("href")
                    full_url = href if href.startswith("http") else self.base_url + href

                    lines = [line.strip() for line in text.splitlines() if line.strip()]
                    program_name = lines[0] if len(lines) > 0 else ""
                    faculty = lines[1].replace('›', ' > ') if len(lines) > 1 else ""
                    university = lines[2] if len(lines) > 2 else ""

                    programs.append({
                        "keyword": keyword,
                        "program_name": program_name,
                        "faculty": faculty,
                        "university": university,
                        "url": full_url,
                        "raw_text": text
                    })
                    print(f"  {idx+1}. {program_name[:50]}...")

                except Exception as e:
                    print(f"  ⚠️ Error processing item #{idx+1}: {e}")

            return programs

        except Exception as e:
            print(f"❌ Error during searching '{keyword}': {e}")
            return []

    async def fetch_program_details(self, page, program):
        """เปิดหน้ารายละเอียดหลักสูตร และดึงข้อมูล ค่าใช้จ่าย และประเภทหลักสูตร"""
        try:
            print(f"📄 Fetching details for: {program['program_name'][:50]}...")
            await page.goto(program["url"], wait_until="networkidle")
            await asyncio.sleep(2)

            details = {
                "keyword": program["keyword"],
                "program_name": program["program_name"],
                "faculty": program["faculty"],
                "university": program["university"],
                "program_type": "ไม่พบข้อมูล",
                "tuition_fee": "ไม่พบข้อมูล",
                "url": program["url"],
                "scrape_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }

            # Selector ที่จะใช้ดึงข้อมูลรายละเอียดที่สนใจ
            type_selectors = [
                "dt:has-text('ประเภทหลักสูตร') + dd",
                ".program-type",
                "[data-field='program_type']",
                "td:has-text('ประเภทหลักสูตร') + td"
            ]
            fee_selectors = [
                "dt:has-text('ค่าใช้จ่าย') + dd",
                "dt:has-text('ค่าธรรมเนียม') + dd",
                ".fee-info",
                ".tuition-fee",
                "[data-field='fee']",
                "td:has-text('ค่าใช้จ่าย') + td"
            ]

            for sel in type_selectors:
                el = await page.query_selector(sel)
                if el:
                    txt = (await el.inner_text()).strip()
                    if txt:
                        details["program_type"] = txt
                        break

            for sel in fee_selectors:
                el = await page.query_selector(sel)
                if el:
                    txt = (await el.inner_text()).strip()
                    if txt:
                        details["tuition_fee"] = txt
                        break

            # หากยังไม่มีข้อมูล ให้ลองดึงจากตาราง
            try:
                rows = await page.query_selector_all("table tr")
                for row in rows:
                    cells = await row.query_selector_all("td, th")
                    if len(cells) >= 2:
                        header = (await cells[0].inner_text()).strip()
                        value = (await cells[1].inner_text()).strip()

                        if "ประเภท" in header and details["program_type"] == "ไม่พบข้อมูล":
                            details["program_type"] = value
                        if any(w in header for w in ["ค่าใช้จ่าย", "ธรรมเนียม", "ค่าเรียน"]) and details["tuition_fee"] == "ไม่พบข้อมูล":
                            details["tuition_fee"] = value
            except:
                pass

            print(f"   ✔️ {details['university'][:30]} - {details['tuition_fee'][:30]}")
            return details

        except Exception as e:
            print(f"   ❌ Failed to fetch details: {e}")
            return None

    async def scrape(self, keywords):
        """จัดการรันกระบวนการสครัปทั้งหมด"""
        results = []
        print("🚀 Starting scraping process...")

        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=True)  # เปลี่ยนเป็น False หากต้องการเห็น browser
            context = await browser.new_context(
                locale="th-TH",
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36"
            )
            page = await context.new_page()

            for kw in keywords:
                programs = await self.perform_search(page, kw)
                if not programs:
                    print(f"❌ ไม่พบหลักสูตรสำหรับคำค้น '{kw}'")
                    continue

                for idx, program in enumerate(programs, 1):
                    print(f"[{idx}/{len(programs)}] Processing program...")
                    detail = await self.fetch_program_details(page, program)
                    if detail:
                        self.collected_data.append(detail)
                    await asyncio.sleep(1.2)  # หน่วงเวลาเพื่อป้องกันการโดนบล็อค

            await browser.close()

        print(f"✅ เสร็จสิ้นการดึงข้อมูลทั้งหมด จำนวนหลักสูตร: {len(self.collected_data)}")
        return self.collected_data

    def save_data(self, filename_prefix="mytcas_scraped"):
        if not self.collected_data:
            print("❌ ไม่มีข้อมูลให้บันทึก")
            return None
        df = pd.DataFrame(self.collected_data)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_file = f"{filename_prefix}_{timestamp}.xlsx"
        csv_file = f"{filename_prefix}_{timestamp}.csv"

        df.to_excel(excel_file, index=False, engine='openpyxl')
        df.to_csv(csv_file, index=False, encoding="utf-8-sig")

        print(f"💾 บันทึกข้อมูลสำเร็จเป็นไฟล์ Excel: {excel_file}")
        print(f"💾 บันทึกข้อมูลสำรองเป็นไฟล์ CSV: {csv_file}")
        return df

async def main():
    print("=== MyTCAS Scraper ===")
    print("เลือกคำค้น:")
    print("1) วิศวกรรม ปัญญาประดิษฐ์")
    print("2) วิศวกรรม คอมพิวเตอร์")
    print("3) ทั้งสองคำ")
    print("4) กำหนดเอง")

    choice = input("เลือก (1-4): ").strip()
    if choice == "1":
        keywords = ["วิศวกรรม ปัญญาประดิษฐ์"]
    elif choice == "2":
        keywords = ["วิศวกรรม คอมพิวเตอร์"]
    elif choice == "3":
        keywords = ["วิศวกรรม ปัญญาประดิษฐ์", "วิศวกรรม คอมพิวเตอร์"]
    elif choice == "4":
        kw_input = input("กรอกคำค้น (คั่นด้วย , ): ")
        keywords = [k.strip() for k in kw_input.split(",") if k.strip()]
    else:
        print("เลือกไม่ถูกต้อง, ใช้ค่าเริ่มต้น 'วิศวกรรม ปัญญาประดิษฐ์'")
        keywords = ["วิศวกรรม ปัญญาประดิษฐ์"]

    print(f"🔎 คำค้นที่ใช้: {keywords}")

    scraper = MyTCASScraper()
    await scraper.scrape(keywords)
    scraper.save_data()

if __name__ == "__main__":
    asyncio.run(main())
