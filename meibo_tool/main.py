"""名簿帳票ツール — エントリーポイント"""

import logging
import os
import sys

# PyInstaller frozen 対応
if getattr(sys, 'frozen', False):
    os.chdir(os.path.dirname(sys.executable))


def main():
    try:
        from gui.app import App
        app = App()
        app.mainloop()
    except Exception:
        logging.exception('アプリの起動に失敗しました')
        try:
            import tkinter.messagebox as _mb
            _mb.showerror('起動エラー', 'アプリの起動に失敗しました。\nログを確認してください。')
        except Exception:
            pass


if __name__ == '__main__':
    main()
