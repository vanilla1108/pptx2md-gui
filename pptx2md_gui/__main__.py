"""pptx2md GUI 的入口点：python -m pptx2md_gui"""

from pptx2md_gui.app import App


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
