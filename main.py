# -*- coding: utf-8 -*-
from multiprocessing import freeze_support
from ui.app import IGKNameApp


def main() -> None:
    app = IGKNameApp()
    app.run()


if __name__ == "__main__":
    freeze_support()
    main()