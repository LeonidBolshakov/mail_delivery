from Frame import *
from loguru import logger

if __name__ == "__main__":
    logger.add("journal.log", format="{time} {level} {message}", level="INFO")

    frame = WidgetOutput()
    frame.mainloop()
