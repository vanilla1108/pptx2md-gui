"""pptx2md GUI 的入口点：python -m pptx2md_gui"""

import multiprocessing as mp

from pptx2md_gui.app import App


def main():
    # Windows 打包后启用 multiprocessing 需要 freeze_support。
    mp.freeze_support()
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
