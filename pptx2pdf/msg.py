#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""利用者に見える名前（プログラム名と，上限・変換器の指定方法）．

このパッケージは単体のコマンド（``pptx2pdf``）としても，**別のツールの一部**としても
動く（md2pptx の ``--pdf`` はここを呼ぶ）．どちらで動いているかで，利用者に見せるべき
名前が変わる：

- 警告や案内の接頭辞は，**利用者が打ったコマンド**の名前でなければならない．
  ``md2pptx deck.md --pdf`` の途中で ``pptx2pdf: warning: …`` と出ても，利用者は
  そんなコマンドを打っていないので何の話か分からない．
- 「上限を延ばすには」「変換器を変えるには」の案内も同じ．このパッケージのフラグは
  ``--timeout`` / ``--converter`` だが，md2pptx から呼ばれているなら
  ``--pdf-timeout`` / ``--pdf-converter`` と案内しなければ通じない．

そこで**メッセージに出す名前だけ**をここに集め，組み込む側が起動時に差し替えられる
ようにする（``set_program_name`` / ``set_hints``）．差し替えないときの既定は
``pptx2pdf`` 自身の名前なので，単体コマンドとしては何もしなくてよい．

変換の挙動には一切関わらない——ここにあるのは表示のための文字列だけ．
参照する側は ``from . import msg`` して ``msg.PROG`` のように**属性で**読むこと
（``from .msg import PROG`` は差し替え前の値を焼き付けてしまう）．
"""
from __future__ import annotations

# 利用者が打ったコマンドの名前．警告の接頭辞と，作業ディレクトリの名前に使う．
PROG = "pptx2pdf"

# 待ちの上限をどう指定するか（打ち切りメッセージに添える）．
HINT_TIMEOUT = "--timeout / PPTX2PDF_TIMEOUT"

# 変換器をどう指定するか（``--converter libreoffice`` のように名前を続けて使うので，
# 環境変数名は含めない）．
HINT_CONVERTER = "--converter"


def set_program_name(name: str) -> None:
    """利用者に見える名前を差し替える（組み込み側が起動時に 1 度だけ呼ぶ）．"""
    global PROG
    PROG = name


def set_hints(*, timeout: str | None = None, converter: str | None = None) -> None:
    """指定方法の案内を差し替える．渡さなかったものは変えない．

    Args:
        timeout: 待ちの上限の指定方法（例 ``"--pdf-timeout / MD2PPTX_PDF_TIMEOUT"``）．
        converter: 変換器の指定に使うフラグ（例 ``"--pdf-converter"``）．
    """
    global HINT_TIMEOUT, HINT_CONVERTER
    if timeout is not None:
        HINT_TIMEOUT = timeout
    if converter is not None:
        HINT_CONVERTER = converter
