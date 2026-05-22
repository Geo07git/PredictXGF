import time
import streamlit as st

# încearcă importul doar local
try:
    import undetected_chromedriver as uc
except Exception:
    uc = None


def get_driver(headless: bool = False):
    # 🔴 dacă suntem în cloud → nu pornim Chrome
    if uc is None or st.runtime.exists():
        return None

    options = uc.ChromeOptions()

    if headless:
        options.add_argument("--headless=new")

    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-popup-blocking")

    driver = uc.Chrome(options=options, version_main=148)

    if not headless:
        time.sleep(1)
        driver.set_window_size(1920, 1080)

    return driver