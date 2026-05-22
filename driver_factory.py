import time
import undetected_chromedriver as uc


def get_driver(headless: bool = False):
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
        #driver.set_window_position(20, 20)

    return driver