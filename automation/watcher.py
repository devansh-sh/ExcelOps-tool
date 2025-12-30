# automation/watcher.py
import time
import os
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from automation.automation_runner import run_automation


WATCH_EXTENSIONS = (".xlsx", ".xls", ".csv")


class ExcelFileHandler(FileSystemEventHandler):
    def on_created(self, event):
        self._handle(event, "CREATED")

    def on_modified(self, event):
        self._handle(event, "MODIFIED")

    def _handle(self, event, event_type):
        if event.is_directory:
            return

        path = event.src_path
        if path.lower().endswith(WATCH_EXTENSIONS):
            print(f"[WATCHER] {event_type}: {os.path.basename(path)}")
            run_automation(path)


def start_watching(folder_path: str):
    folder_path = os.path.abspath(folder_path)

    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

    print("=" * 50)
    print("[WATCHER] STARTED")
    print("[WATCHER] Folder:", folder_path)
    print("Drop Excel files into this folder")
    print("Press Ctrl+C to stop")
    print("=" * 50)

    event_handler = ExcelFileHandler()
    observer = Observer()
    observer.schedule(event_handler, folder_path, recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n[WATCHER] STOPPED")
        observer.stop()

    observer.join()


if __name__ == "__main__":
    # WATCH THIS EXACT FOLDER
    WATCH_FOLDER = "./watch_folder"
    start_watching(WATCH_FOLDER)
