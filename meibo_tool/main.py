"""名簿帳票ツール — エントリーポイント"""

import sys
import os

# PyInstaller frozen 対応
if getattr(sys, 'frozen', False):
    os.chdir(os.path.dirname(sys.executable))


def main():
    from gui.app import App
    app = App()
    app.mainloop()


if __name__ == '__main__':
    main()
