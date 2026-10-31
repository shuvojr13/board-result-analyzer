"""
scraper.py - Playwright Web Automation for Bangladesh Education Board Results
Handles headless browser automation, math captcha solving, and data extraction.
"""

import asyncio
import logging
import random
import re
import sys
from typing import Any, Callable, Dict, List, Optional
import requests
import urllib3
from bs4 import BeautifulSoup
from playwright.async_api import Browser, BrowserContext, Page, async_playwright

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

if sys.platform == "win32":
    try:
        asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
    except Exception:
        pass

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

PORTAL_URL = "http://www.educationboardresults.gov.bd/"

# Board mapping helper
BOARD_MAP = {
    "dhaka": "dhaka",
    "barisal": "barisal",
    "chittagong": "chittagong",
    "comilla": "comilla",
    "dinajpur": "dinajpur",
    "jessore": "jessore",
    "mymensingh": "mymensingh",
    "rajshahi": "rajshahi",
    "sylhet": "sylhet",
    "madrasah": "madrasah",
    "technical": "tec",
    "tec": "tec",
    "dibs": "dibs"
}

# Exam mapping helper
EXAM_MAP = {
    "ssc": "ssc",
    "ssc_voc": "ssc_voc",
    "hsc": "hsc",
    "hsc_voc": "hsc_voc",
    "hsc_hbm": "hsc_hbm",
    "dakhil": "dakhil",
    "jsc": "jsc"
}


def solve_math_captcha_from_html(html_text: str) -> Optional[int]:
    """
    Finds arithmetic expression like '4 + 7' or '8+3' in the captcha region.
    """
    # 1. Regex search for arithmetic expression near value_s
    # Look for patterns like "5 + 4" or " 7 + 8 "
    patterns = [
        r'(\d+)\s*\+\s*(\d+)',
        r'(\d+)\s*plus\s*(\d+)',
    ]
    
    # Try finding expression near value_s input first
    soup = BeautifulSoup(html_text, "html5lib")
    value_s = soup.find("input", {"name": "value_s"})
    if value_s:
        parent_tr = value_s.find_parent("tr")
        if parent_tr:
            tr_text = parent_tr.get_text()
            match = re.search(r'(\d+)\s*\+\s*(\d+)', tr_text)
            if match:
                a, b = int(match.group(1)), int(match.group(2))
                logger.info(f"Solved Captcha from parent row: {a} + {b} = {a + b}")
                return a + b

    # Fallback: scan full HTML for digits with plus sign
    matches = re.findall(r'(\d+)\s*\+\s*(\d+)', html_text)
    if matches:
        # Take the last or first match
        a, b = int(matches[0][0]), int(matches[0][1])
        logger.info(f"Solved Captcha from general HTML: {a} + {b} = {a + b}")
        return a + b

    return None


class ResultScraper:
    def __init__(self, headless: bool = True, timeout: int = 30000):
        self.headless = headless
        self.timeout = timeout
        self.playwright = None
        self.browser: Optional[Browser] = None
        self.context: Optional[BrowserContext] = None

    async def start(self):
        """Initializes Playwright browser session."""
        self.playwright = await async_playwright().start()
        self.browser = await self.playwright.chromium.launch(
            headless=self.headless,
            args=[
                "--no-sandbox",
                "--disable-setuid-sandbox",
                "--disable-dev-shm-usage",
                "--disable-blink-features=AutomationControlled",
            ]
        )
        self.context = await self.browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/123.0.0.0 Safari/537.36"
            ),
            viewport={"width": 1280, "height": 800}
        )

    async def stop(self):
        """Closes browser session."""
        try:
            if self.context:
                await self.context.close()
            if self.browser:
                await self.browser.close()
            if self.playwright:
                await self.playwright.stop()
        except Exception as e:
            logger.warning(f"Error closing Playwright browser: {e}")

    async def fetch_student_result(
        self,
        exam: str,
        year: str,
        board: str,
        roll: str,
        reg: str,
        max_retries: int = 3
    ) -> Dict[str, Any]:
        """
        Navigates to the portal, fills form, solves captcha, and parses student result.
        Retries up to max_retries on captcha error or temporary network glitch.
        """
        board_val = BOARD_MAP.get(board.lower(), board.lower())
        exam_val = EXAM_MAP.get(exam.lower(), exam.lower())

        for attempt in range(1, max_retries + 1):
            page: Optional[Page] = None
            alert_message = None

            try:
                page = await self.context.new_page()

                # Handle JavaScript alert popups (e.g. wrong captcha or result not found)
                async def on_dialog(dialog):
                    nonlocal alert_message
                    alert_message = dialog.message
                    logger.warning(f"Browser Dialog Alert: {alert_message}")
                    await dialog.accept()

                page.on("dialog", on_dialog)

                logger.info(f"Navigating to {PORTAL_URL} (Attempt {attempt}/{max_retries})")
                await page.goto(PORTAL_URL, wait_until="domcontentloaded", timeout=self.timeout)

                # Wait for form elements
                await page.wait_for_selector('input[name="roll"], #roll', timeout=self.timeout)

                # Select Examination
                exam_selector = 'select[name="sr"], select[name="exam"], select[name="examination"], #sr, #examination'
                if await page.query_selector(exam_selector):
                    try:
                        await page.select_option(exam_selector, value=exam_val)
                    except Exception:
                        # Try selecting by label/text if value mismatch
                        await page.select_option(exam_selector, label=exam.upper())

                # Select Year
                year_selector = 'select[name="year"], #year'
                try:
                    await page.select_option(year_selector, value=str(year), timeout=5000)
                except Exception:
                    await page.close()
                    return {
                        "Roll": roll,
                        "Registration": reg,
                        "Exam": exam.upper(),
                        "Year": year,
                        "Board": board.title(),
                        "Name": "N/A",
                        "Result": "Year Not Published",
                        "GPA": "0.00",
                        "Status": f"Results for year {year} are not yet published or available on the official board website. Turn on 'Test / Demo Mode' to preview.",
                        "Subject_Grades": {},
                        "Subject_Grades_Summary": f"Results for {year} are not available on the official website."
                    }

                # Select Board
                board_selector = 'select[name="board"], #board'
                try:
                    await page.select_option(board_selector, value=board_val)
                except Exception:
                    await page.select_option(board_selector, label=board.capitalize())

                # Enter Roll
                roll_selector = 'input[name="roll"], #roll'
                await page.fill(roll_selector, str(roll).strip())

                # Enter Registration (if provided)
                reg_selector = 'input[name="reg"], #reg'
                if reg:
                    await page.fill(reg_selector, str(reg).strip())

                # Captcha Handling: Extract and solve math problem
                html_content = await page.content()
                captcha_answer = solve_math_captcha_from_html(html_content)

                if captcha_answer is None:
                    # Alternative: directly evaluate text in the element before #value_s
                    try:
                        math_text = await page.locator("tr:has(input[name='value_s']) td:first-child").inner_text()
                        m = re.search(r'(\d+)\s*\+\s*(\d+)', math_text)
                        if m:
                            captcha_answer = int(m.group(1)) + int(m.group(2))
                    except Exception:
                        pass

                if captcha_answer is None:
                    logger.error("Could not find or solve captcha. Retrying...")
                    await page.close()
                    await asyncio.sleep(1)
                    continue

                captcha_selector = 'input[name="value_s"], #value_s'
                await page.fill(captcha_selector, str(captcha_answer))

                # Submit the form
                submit_btn = 'input[type="submit"], input[name="button2"], #button2'
                await asyncio.gather(
                    page.wait_for_load_state("domcontentloaded", timeout=self.timeout),
                    page.click(submit_btn)
                )

                # Give the page a moment to settle or trigger alert
                await asyncio.sleep(1)

                if alert_message:
                    # Captcha was incorrect or student record not found
                    if "captcha" in alert_message.lower():
                        logger.warning(f"Captcha rejected by server for Roll {roll}. Retrying...")
                        await page.close()
                        continue
                    else:
                        logger.info(f"Server returned alert: '{alert_message}' for Roll: {roll}")
                        await page.close()
                        return {
                            "Roll": roll,
                            "Registration": reg,
                            "Exam": exam.upper(),
                            "Year": year,
                            "Board": board.title(),
                            "Name": "N/A",
                            "Result": "Not Found",
                            "GPA": "0.00",
                            "Status": alert_message,
                            "Subject_Grades": {},
                            "Subject_Grades_Summary": alert_message
                        }

                # Parse the response page
                page_html = await page.content()
                parsed_data = self._parse_result_html(page_html, roll, reg, exam, year, board)

                await page.close()
                return parsed_data

            except Exception as e:
                logger.warning(f"Error scraping Roll {roll} (Attempt {attempt}): {e}")
                if page:
                    try:
                        await page.close()
                    except Exception:
                        pass
                if attempt < max_retries:
                    await asyncio.sleep(1.5)

        # Fallback if all retries failed
        return {
            "Roll": roll,
            "Registration": reg,
            "Exam": exam.upper(),
            "Year": year,
            "Board": board.title(),
            "Name": "N/A",
            "Result": "Failed / Error",
            "GPA": "0.00",
            "Status": "Connection / Captcha Timeout",
            "Subject_Grades": {},
            "Subject_Grades_Summary": "Connection / Captcha Timeout"
        }

    def _parse_result_html(
        self,
        html_str: str,
        roll: str,
        reg: str,
        exam: str,
        year: str,
        board: str
    ) -> Dict[str, Any]:
        """
        Parses the result page HTML tables to extract student details, GPA, and subject grades.
        """
        soup = BeautifulSoup(html_str, "html5lib")

        # Check for error / not found message in page body
        text_content = soup.get_text()
        if "result not found" in text_content.lower() or "no result found" in text_content.lower():
            return {
                "Roll": roll,
                "Registration": reg,
                "Exam": exam.upper(),
                "Year": year,
                "Board": board.title(),
                "Name": "N/A",
                "Result": "Result Not Found",
                "GPA": "0.00",
                "Status": "Not Found",
                "Subject_Grades": {},
                "Subject_Grades_Summary": "Result Not Found"
            }

        student_info: Dict[str, str] = {}
        subject_grades: Dict[str, str] = {}

        # Scan all tables for student details and subject breakdown
        tables = soup.find_all("table")

        for table in tables:
            rows = table.find_all("tr")
            for row in rows:
                cols = [td.get_text(strip=True) for td in row.find_all(["td", "th"])]
                
                # Check 2-column or 4-column key-value pairs (e.g. Roll No, Name, GPA)
                if len(cols) == 2:
                    k, v = cols[0].rstrip(":").strip(), cols[1].strip()
                    if k:
                        student_info[k.lower()] = v
                elif len(cols) == 4:
                    k1, v1 = cols[0].rstrip(":").strip(), cols[1].strip()
                    k2, v2 = cols[2].rstrip(":").strip(), cols[3].strip()
                    if k1:
                        student_info[k1.lower()] = v1
                    if k2:
                        student_info[k2.lower()] = v2

                # Check 3-column subject grade rows (Code, Subject Name, Grade)
                elif len(cols) == 3:
                    c1, c2, c3 = cols[0], cols[1], cols[2]
                    # Check if c3 looks like a grade (A+, A, A-, B, C, D, F)
                    if c3 in ["A+", "A", "A-", "B", "C", "D", "F"]:
                        subj_name = f"{c1} {c2}".strip() if c1.isdigit() else c2
                        subject_grades[subj_name] = c3

        # Match common field variants
        name = (
            student_info.get("name")
            or student_info.get("student's name")
            or student_info.get("student name")
            or "N/A"
        )
        father_name = student_info.get("father's name") or student_info.get("father name") or ""
        mother_name = student_info.get("mother's name") or student_info.get("mother name") or ""
        gpa = student_info.get("gpa") or student_info.get("gpa (out of 5)") or "0.00"
        result_status = student_info.get("result") or student_info.get("result status") or ("PASSED" if gpa and gpa != "0.00" else "FAILED")
        group = student_info.get("group") or student_info.get("stream") or ""
        type_candidate = (
            student_info.get("type")
            or student_info.get("candidate type")
            or student_info.get("student type")
            or ""
        ).strip().upper()
        if not type_candidate or type_candidate in ["NAN", "NONE", "-"]:
            if 0 < len(subject_grades) <= 4:
                type_candidate = "IRREGULAR"
            else:
                type_candidate = "REGULAR"
        elif "IRREG" in type_candidate:
            type_candidate = "IRREGULAR"
        elif "REG" in type_candidate:
            type_candidate = "REGULAR"
        elif "IMP" in type_candidate:
            type_candidate = "IMPROVEMENT"
        elif "PRIV" in type_candidate:
            type_candidate = "PRIVATE"
        institute = student_info.get("institute") or student_info.get("institution") or ""

        # Subject Grades summary string: "Bangla: A+, English: A, Mathematics: A+"
        grades_summary = ", ".join([f"{k}: {v}" for k, v in subject_grades.items()]) if subject_grades else "N/A"

        return {
            "Roll": roll,
            "Registration": reg,
            "Exam": exam.upper(),
            "Year": year,
            "Board": board.title(),
            "Name": name,
            "Father Name": father_name,
            "Mother Name": mother_name,
            "Institute": institute,
            "Group": group,
            "Type": type_candidate,
            "Result": result_status,
            "GPA": gpa,
            "Status": "Success" if name != "N/A" else "Result Found",
            "Subject_Grades": subject_grades,
            "Subject_Grades_Summary": grades_summary
        }


# ==========================================
# Simulated Demo Scraper (For Instant Testing)
# ==========================================
async def generate_simulated_result(
    roll: str,
    reg: str,
    exam: str,
    year: str,
    board: str
) -> Dict[str, Any]:
    """
    Generates a realistic mock result for testing and demonstration
    when live board servers are unreachable or for local dev.
    """
    await asyncio.sleep(random.uniform(0.6, 1.2))  # Realistic network latency simulation

    sample_names = [
        "MD. TANVIR HASAN", "NUSRAT JAHAN", "RAKIBUL ISLAM", "FATEMA AKTER",
        "SAKIB AL HASAN", "SADIA SULTANA", "ARIFUR RAHMAN", "MIM AKTER",
        "MEHEDI HASAN", "JANNATUL MAOWA", "KAMRUN NAHAR", "DIPA AKTER"
    ]

    # Diverse groups and realistic curriculum
    group_curricula = {
        "SCIENCE": [
            ("101", "BANGLA"), ("107", "ENGLISH"), ("275", "INFORMATION & COMMUNICATION TECHNOLOGY"),
            ("174", "PHYSICS"), ("175", "CHEMISTRY"), ("176", "BIOLOGY"), ("129", "HIGHER MATHEMATICS")
        ],
        "BUSINESS STUDIES": [
            ("101", "BANGLA"), ("107", "ENGLISH"), ("275", "INFORMATION & COMMUNICATION TECHNOLOGY"),
            ("253", "ACCOUNTING"), ("277", "BUSINESS ORGANIZATION & MANAGEMENT"),
            ("286", "PRODUCTION MANAGEMENT & MARKETING"), ("292", "FINANCE, BANKING & INSURANCE")
        ],
        "HUMANITIES": [
            ("101", "BANGLA"), ("107", "ENGLISH"), ("275", "INFORMATION & COMMUNICATION TECHNOLOGY"),
            ("109", "ECONOMICS"), ("267", "ISLAMIC HISTORY & CULTURE"),
            ("269", "CIVICS & GOOD GOVERNANCE"), ("271", "SOCIAL WORK")
        ]
    }

    group = random.choice(["SCIENCE", "BUSINESS STUDIES", "HUMANITIES"])
    subjects = group_curricula[group]
    passing_pool = ["A+", "A", "A-", "B", "C", "D"]

    # Determine outcome: 4% Absent, 62% Pass, 19% 1-subject fail, 10% 2-subject fail, 5% 3-subject fail
    outcome_roll = random.random()
    subject_grades = {}

    if outcome_roll < 0.04:
        # Absent candidate
        selected_gpa = "0.00"
        result_status = "ABSENT"
        subject_grades = {}
    elif outcome_roll < 0.66:
        # Passed
        selected_gpa = random.choice(["5.00", "4.89", "4.72", "4.50", "4.17", "3.83", "3.50", "3.17"])
        for code, name in subjects:
            subject_grades[f"{code} {name}"] = random.choice(passing_pool)
        result_status = "PASSED"
    elif outcome_roll < 0.85:
        # 1-Subject Fail
        selected_gpa = "0.00"
        result_status = "FAIL"
        failed_idx = random.randrange(len(subjects))
        for i, (code, name) in enumerate(subjects):
            if i == failed_idx:
                subject_grades[f"{code} {name}"] = "F"
            else:
                subject_grades[f"{code} {name}"] = random.choice(passing_pool)
    elif outcome_roll < 0.95:
        # 2-Subject Fail
        selected_gpa = "0.00"
        result_status = "FAIL"
        failed_indices = set(random.sample(range(len(subjects)), 2))
        for i, (code, name) in enumerate(subjects):
            if i in failed_indices:
                subject_grades[f"{code} {name}"] = "F"
            else:
                subject_grades[f"{code} {name}"] = random.choice(passing_pool)
    else:
        # 3-Subject Fail
        selected_gpa = "0.00"
        result_status = "FAIL"
        failed_indices = set(random.sample(range(len(subjects)), 3))
        for i, (code, name) in enumerate(subjects):
            if i in failed_indices:
                subject_grades[f"{code} {name}"] = "F"
            else:
                subject_grades[f"{code} {name}"] = random.choice(passing_pool)

    summary = ", ".join([f"{k}: {v}" for k, v in subject_grades.items()]) if subject_grades else "ABSENT"

    candidate_type = "REGULAR" if random.random() < 0.75 else "IRREGULAR"

    return {
        "Roll": roll,
        "Registration": reg,
        "Exam": exam.upper(),
        "Year": year,
        "Board": board.title(),
        "Name": random.choice(sample_names),
        "Father Name": "MD. ABDUR RAHIM",
        "Mother Name": "MRS. ROKEYA BEGUM",
        "Group": group,
        "Type": candidate_type,
        "Result": result_status,
        "GPA": selected_gpa,
        "Status": "Success (Simulated)",
        "Subject_Grades": subject_grades,
        "Subject_Grades_Summary": summary
    }


def fetch_via_http_session(exam: str, year: str, board: str, roll: str, reg: str) -> Optional[Dict[str, Any]]:
    """Fast HTTP fallback using requests without browser overhead."""
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36"
    }
    board_val = BOARD_MAP.get(board.lower(), board.lower())
    exam_val = EXAM_MAP.get(exam.lower(), exam.lower())
    try:
        s = requests.Session()
        r = s.get("http://www.educationboardresults.gov.bd/", headers=headers, timeout=8)
        if r.status_code == 200:
            captcha_ans = solve_math_captcha_from_html(r.text)
            if captcha_ans is not None:
                payload = {
                    "sr": "3",
                    "et": "2",
                    "exam": exam_val,
                    "year": str(year),
                    "board": board_val,
                    "roll": str(roll).strip(),
                    "reg": str(reg).strip(),
                    "value_s": str(captcha_ans),
                    "button2": "Submit"
                }
                res_post = s.post("http://www.educationboardresults.gov.bd/result.php", data=payload, headers=headers, timeout=8)
                if res_post.status_code == 200:
                    dummy = ResultScraper()
                    return dummy._parse_result_html(res_post.text, roll, reg, exam, year, board)
    except Exception as e:
        logger.warning(f"HTTP fallback error: {e}")
    return None


CUMILLA_HSC_PORTAL_URL = "https://hscresult.comillaboard.gov.bd/result.php"


def fetch_cumilla_hsc_result_sync(
    roll: str,
    year: str = "",
    reg: str = ""
) -> Dict[str, Any]:
    """
    Directly fetches HSC result from the dedicated BISE Cumilla server:
    https://hscresult.comillaboard.gov.bd/result.php
    Requires ONLY Roll number (Registration and Captcha are not needed).
    """
    clean_roll = str(roll).strip()
    clean_reg = str(reg).strip() if reg else ""
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/123.0.0.0 Safari/537.36"
        ),
        "Referer": "https://hscresult.comillaboard.gov.bd/"
    }

    try:
        resp = requests.post(
            CUMILLA_HSC_PORTAL_URL,
            data={"roll": clean_roll, "button2": "Submit"},
            headers=headers,
            verify=False,
            timeout=12
        )

        if resp.status_code == 200:
            scraper = ResultScraper()
            m = re.search(r'HSC Result\s*(\d{4})', resp.text, re.IGNORECASE)
            detected_year = m.group(1) if m else (str(year) if year else "2026")

            result = scraper._parse_result_html(
                html_str=resp.text,
                roll=clean_roll,
                reg=clean_reg,
                exam="HSC",
                year=detected_year,
                board="Cumilla"
            )
            return result
        else:
            return {
                "Roll": clean_roll,
                "Registration": clean_reg,
                "Exam": "HSC",
                "Year": str(year) if year else "2026",
                "Board": "Cumilla",
                "Name": "N/A",
                "Result": f"Server Error ({resp.status_code})",
                "GPA": "0.00",
                "Status": f"Cumilla Board server returned HTTP {resp.status_code}",
                "Subject_Grades": {},
                "Subject_Grades_Summary": f"Server responded with error status {resp.status_code}"
            }
    except Exception as e:
        logger.error(f"Error fetching Cumilla HSC result for roll {clean_roll}: {e}")
        return {
            "Roll": clean_roll,
            "Registration": clean_reg,
            "Exam": "HSC",
            "Year": str(year) if year else "2026",
            "Board": "Cumilla",
            "Name": "N/A",
            "Result": "Network Error",
            "GPA": "0.00",
            "Status": f"Failed to connect to {CUMILLA_HSC_PORTAL_URL}: {str(e)}",
            "Subject_Grades": {},
            "Subject_Grades_Summary": "Failed to connect to Cumilla board server"
        }


async def fetch_cumilla_hsc_result(
    roll: str,
    year: str = "",
    reg: str = ""
) -> Dict[str, Any]:
    """Async wrapper for fetch_cumilla_hsc_result_sync."""
    return await asyncio.to_thread(fetch_cumilla_hsc_result_sync, roll, year, reg)


async def scrape_single_student(
    exam: str,
    year: str,
    board: str,
    roll: str,
    reg: str = "",
    demo_mode: bool = False
) -> Dict[str, Any]:
    """
    Fetches the result for an individual student.
    Uses simulated data if demo_mode is True, otherwise launches Playwright with HTTP fallback.
    """
    if demo_mode:
        logger.info(f"Checking individual Roll: {roll} in DEMO mode.")
        return await generate_simulated_result(roll, reg, exam, year, board)

    # Special Fast Dedicated Server for Cumilla Board HSC Results (Only Roll required)
    if board.lower() in ["comilla", "cumilla"] and "hsc" in exam.lower():
        logger.info(f"Fetching Cumilla HSC result for Roll: {roll} via https://hscresult.comillaboard.gov.bd/")
        return await fetch_cumilla_hsc_result(roll=roll, year=year, reg=reg)

    scraper = ResultScraper(headless=True)
    try:
        await scraper.start()
        logger.info(f"Live Playwright check for individual Roll: {roll}, Reg: {reg}")
        return await scraper.fetch_student_result(exam, year, board, roll, reg)
    except Exception as e:
        logger.warning(f"Playwright error: {e}. Trying HTTP fallback...")
        http_result = await asyncio.to_thread(fetch_via_http_session, exam, year, board, roll, reg)
        if http_result:
            return http_result

        # Return structured fallback instead of raising raw exception
        return {
            "Roll": roll,
            "Registration": reg,
            "Exam": exam.upper(),
            "Year": year,
            "Board": board.title(),
            "Name": "N/A",
            "Result": "Connection Error",
            "GPA": "0.00",
            "Status": "Unable to connect to the education board server. The government site may be slow or temporarily down. Please try again or switch on 'Test / Demo Mode'.",
            "Subject_Grades": {},
            "Subject_Grades_Summary": "Connection to official portal timed out."
        }
    finally:
        await scraper.stop()


async def scrape_student_batch(
    student_records: List[Dict[str, str]],
    exam: str,
    year: str,
    board: str,
    demo_mode: bool = False,
    on_progress: Optional[Callable[[int, int, Dict[str, Any]], Any]] = None
) -> List[Dict[str, Any]]:
    """
    Iterates through a list of student records, fetching results for each one.
    Calls `on_progress(current_index, total_count, student_result)` for real-time reporting.
    """
    results = []
    total = len(student_records)

    if demo_mode:
        logger.info(f"Running batch of {total} students in DEMO mode.")
        for idx, item in enumerate(student_records, start=1):
            roll = str(item.get("Roll", "")).strip()
            reg = str(item.get("Registration", "")).strip()
            result = await generate_simulated_result(roll, reg, exam, year, board)
            results.append(result)
            if on_progress:
                await on_progress(idx, total, result)
        return results

    # Special Fast Dedicated Server for Cumilla Board HSC Results (Only Roll required)
    if board.lower() in ["comilla", "cumilla"] and "hsc" in exam.lower():
        logger.info(f"Running fast batch extraction for {total} Cumilla HSC students via https://hscresult.comillaboard.gov.bd/")
        for idx, item in enumerate(student_records, start=1):
            roll = str(item.get("Roll", "")).strip()
            reg = str(item.get("Registration", "")).strip()
            logger.info(f"[{idx}/{total}] Fetching Cumilla HSC Roll: {roll}")

            res = await fetch_cumilla_hsc_result(roll=roll, year=year, reg=reg)
            results.append(res)

            if on_progress:
                await on_progress(idx, total, res)

            # Modest delay between server requests (150ms) to ensure stability
            if idx < total:
                await asyncio.sleep(0.15)

        return results

    # Live Playwright Scraping
    scraper = ResultScraper(headless=True)
    try:
        await scraper.start()
        for idx, item in enumerate(student_records, start=1):
            roll = str(item.get("Roll", "")).strip()
            reg = str(item.get("Registration", "")).strip()
            logger.info(f"[{idx}/{total}] Scraping Roll: {roll}, Reg: {reg}")

            res = await scraper.fetch_student_result(exam, year, board, roll, reg)
            results.append(res)

            if on_progress:
                await on_progress(idx, total, res)

            # Politeness delay between server requests (1 second)
            if idx < total:
                await asyncio.sleep(1.0)

    finally:
        await scraper.stop()

    return results
