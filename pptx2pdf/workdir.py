#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""使い捨ての作業ディレクトリの片付け（md2pptx Issue #58 から）．

成果物をアトミックに差し替えるために，出力先の隣へ使い捨ての作業ディレクトリを作って
そこで組み立てる（PDF は ``converter.convert``）．ほかにも LibreOffice の使い捨て
プロファイルや PowerPoint コンテナ内の staging など，**「自分で作って自分で捨てる
作業ディレクトリ」**はいくつもある．その片付け方をここに 1 つだけ置く．

規則は 2 つで，どちらも外すと運用が壊れる．

- **片付けの失敗で処理の成否を変えない．** 片付けに入る時点で本来の仕事（保存・変換）は
  終わっている．そこで例外を投げると，成功した実行が失敗になり，しかも本体が投げた例外が
  あればそれを握りつぶして置き換えてしまう．
- **それでも黙って残さない．** 消せなかったことを誰にも伝えないと，``--watch`` のような
  作り直しの運用では保存のたびに 1 つずつ溜まっていく．しかも出力先ディレクトリに溜まる
  ものは利用者の目に触れる．原因は環境側（Windows で走査ソフトがファイルを掴んでいる等）
  なので**ここで再試行はせず**，起きた事実だけを伝える．

``tempfile.TemporaryDirectory`` に置き換えてはいけない——既定（``ignore_cleanup_errors=
False``）では**片付けの失敗で例外を投げる**ので，1 つめの規則が崩れる．

メッセージは利用者に見える文字列なので，写しを増やさない意味でもここが唯一の出どころ．
外部依存は持たない（このパッケージを組み込む側からも使える）．
"""
from __future__ import annotations

import os
import shutil
import sys
import tempfile

from . import msg


def create(where: str | None = None, prefix: str | None = None) -> str:
    """作業ディレクトリを作ってパスを返す．

    Args:
        where: 作る場所．``None`` ならシステムの一時領域．成果物を組み立てるときは
            **出力先と同じディレクトリ**を渡す——同一ファイルシステム上に置くことで，
            仕上げの ``os.replace`` が必ずアトミックになる（EXDEV が起こりえない）．
        prefix: 名前の接頭辞．``None`` なら ``.<プログラム名>-``（例 ``.pptx2pdf-``）．
            ドット始まりなのは，**出力先に作る**もの（既定の使い方）を利用者の目に
            付かせないため．名前が利用者に見える以上，接頭辞は**利用者が打ったコマンド
            の名前**から作る（``msg.PROG``）．システムの一時領域へ作るときは隠す理由が
            無いので，``prefix=f"{msg.PROG}-lo-"`` のように明示して渡す．

    Raises:
        OSError: 作れなかったとき．**ここでは整形しない**——「作業場所を用意できなかった」
            ことをどう伝えるかは呼び出し側の事情で違うため（``convert`` は ``PdfError``
            にして終了コードを変えず，組み込み側は何をしようとして失敗したかを添えて
            送出する）．
    """
    if prefix is None:
        prefix = f".{msg.PROG}-"
    return tempfile.mkdtemp(dir=where, prefix=prefix)


def discard(work: str) -> None:
    """作業ディレクトリを捨てる．**例外は投げない**（モジュール docstring の規則）．

    消せなければ stderr に 1 行出すだけで、処理の成否は変えない．
    """
    # ignore_errors=True なのは，エラーが起きても走査を続けて**消せるものは消す**ため．
    # 外すと最初のエラーで止まり，残りが丸ごと残る．
    shutil.rmtree(work, ignore_errors=True)
    if os.path.isdir(work):
        sys.stderr.write(f"{msg.PROG}: warning: could not remove {work}\n")
