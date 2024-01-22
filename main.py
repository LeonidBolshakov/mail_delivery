from Frame import *
from loguru import logger


def main():
    logger.add("journal.log", format="{time} {level} {message}", level="INFO")
    root = Tk()

    Frame(root)

    root.mainloop()


if __name__ == "__main__":
    main()
