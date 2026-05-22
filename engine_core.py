import time
import random
import threading

print("USING engine_core.py with message_queue")


class ScraperEngine:
    def __init__(self, driver_factory, extractor, message_queue):
        self.driver_factory = driver_factory
        self.extractor = extractor
        self.message_queue = message_queue
        self.queue = []
        self.results = {}
        self.running = False
        self.thread = None

    def load_queue(self, pairs):
        self.queue = pairs

    def start(self):
        if self.running:
            return
        self.running = True
        self.thread = threading.Thread(target=self._run, daemon=True)
        self.thread.start()

    def stop(self):
        self.running = False
        self.log("⛔ STOP requested")

    def log(self, msg):
        self.message_queue.put(str(msg))

    def _run(self):
        driver = None
        try:
            self.log("🚀 ENGINE STARTED")
            self.log("🔧 Creating driver...")
            driver = self.driver_factory()
            self.log("✅ Driver created")

            total = len(self.queue)
            self.log(f"📦 Queue size: {total}")

            for i, (league, url) in enumerate(self.queue):
                if not self.running:
                    self.log("🛑 Stop flag detected")
                    break

                self.log(f"[{i+1}/{total}] {league}")
                self.message_queue.put(("__PROGRESS__", i / max(total, 1)))

                success = False
                for attempt in range(2):
                    if not self.running:
                        break

                    self.log(f"🔁 Încercare {attempt+1}/2")
                    try:
                        df = self.extractor(driver, url, self.log)
                        if df is not None:
                            self.results[league] = df
                            self.log(f"✅ OK: {league}")
                            success = True
                            break
                        else:
                            self.log(f"⚠️ Fără date: {league}")
                    except Exception as e:
                        self.log(f"💥 Extractor error: {type(e).__name__}: {e}")

                    time.sleep(random.uniform(2, 4))

                if not success:
                    self.log(f"❌ FAILED: {league}")

            self.message_queue.put(("__PROGRESS__", 1.0))

        except Exception as e:
            self.log(f"💥 ENGINE FATAL: {type(e).__name__}: {e}")
        finally:
            if driver is not None:
                try:
                    driver.quit()
                except Exception as e:
                    self.log(f"⚠️ driver.quit error: {type(e).__name__}: {e}")

            self.running = False
            self.log("🏁 ENGINE STOPPED")