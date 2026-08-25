# -*- coding: utf-8 -*-
"""pptx2pdf — pptx を PDF に変換する（PowerPoint / LibreOffice / 任意のコマンド）．

コマンド ``pptx2pdf in.pptx`` として使うほか，ライブラリとしても使える::

    import pptx2pdf
    pptx2pdf.convert("deck.pptx", "deck.pdf", None)     # None = auto

組み込む側（別のコマンドの一部として動かす場合）は，利用者に見える名前を
``set_program_name`` / ``set_hints`` で自分のものに差し替えられる（``msg`` 参照）．

もとは md2pptx の PDF 変換部分（``md2pptx/pdf.py`` + ``md2pptx/workdir.py``）で，
python-pptx に依存しないので単体のコマンドとして切り出した．
"""
from __future__ import annotations

from . import workdir
from .converter import (
    ENV_CONVERTER,
    ENV_TIMEOUT,
    PdfError,
    convert,
    default_pdf_path,
)
from .msg import set_hints, set_program_name

__version__ = "0.1.0"

__all__ = [
    "ENV_CONVERTER",
    "ENV_TIMEOUT",
    "PdfError",
    "__version__",
    "convert",
    "default_pdf_path",
    "set_hints",
    "set_program_name",
    "workdir",
]
