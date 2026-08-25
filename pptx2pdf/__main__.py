# -*- coding: utf-8 -*-
"""``python3 -m pptx2pdf`` の入口．

このモジュールは ``-m`` 実行時にのみ ``__main__`` として読み込まれるため，
``if __name__ == "__main__":`` のガードは不要（常に main を呼ぶ）．
"""
from .cli import main

main()
