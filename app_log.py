"""
app_log.py - ведение журнала программы.
Пишет в protocolbot.log рядом с программой.
Файл можно удалять в любой момент - он создастся заново.
"""
import os
import sys
import logging
import traceback

LOG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "protocolbot.log")

logging.basicConfig(
    filename=LOG_PATH,
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    encoding="utf-8",
)


def install():
    """Вешает крючки: окна ошибок -> в лог; необработанные ошибки -> в лог с traceback."""
    from tkinter import messagebox

    orig_showerror = messagebox.showerror

    def showerror_with_log(*args, **kwargs):
        try:
            logging.error("ОКНО ОШИБКИ: " + " | ".join(str(a) for a in args))
        except Exception:
            pass
        return orig_showerror(*args, **kwargs)

    messagebox.showerror = showerror_with_log

    def log_exception(exc_type, exc_value, exc_tb):
        logging.error(
            "НЕОБРАБОТАННАЯ ОШИБКА:\n"
            + "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
        )
        try:
            orig_showerror(
                "Ошибка",
                "Произошла непредвиденная ошибка.\n"
                "Подробности записаны в protocolbot.log - пришлите этот файл для разбора.\n\n"
                + str(exc_value),
            )
        except Exception:
            pass

    sys.excepthook = log_exception
    logging.info("=== Программа запущена ===")