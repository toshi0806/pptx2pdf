#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""``pptx2pdf`` コマンドの入口．

使い方::

    pptx2pdf deck.pptx                      # deck.pdf を隣に作る
    pptx2pdf deck.pptx -o /tmp/out.pdf      # 出力先を指定
    pptx2pdf deck.pptx --converter libreoffice
    pptx2pdf deck.pptx --converter 'unoconvert {input} {output}'

変換の中身は ``converter`` にある．ここでやるのは引数と環境変数の解決，成否の報告，
終了コードだけ．失敗は ``pptx2pdf: <理由>`` の 1 行で伝えて終了コード 1 にする
（成果物である PDF ができていないので，md2pptx の ``--pdf``（pptx は保存できている）
とは違い**失敗として終わる**）．
"""
from __future__ import annotations

import argparse
import os
import sys

from . import __version__
from .converter import ENV_CONVERTER, ENV_TIMEOUT, PdfError, convert, default_pdf_path


def _build_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(
        prog="pptx2pdf",
        description="pptx を PDF に変換する（PowerPoint / LibreOffice / 任意のコマンド）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "変換器（--converter）:\n"
            "  auto         既定．native PowerPoint → LibreOffice の順に使えるものを選ぶ．\n"
            "               選んだ変換器が失敗しても次へは落とさない（忠実度が黙って\n"
            "               入れ替わらないように）．\n"
            "  powerpoint   実 PowerPoint（macOS: AppleScript / Windows: COM）．\n"
            "  libreoffice  soffice --headless．\n"
            "  コマンド行    {input} / {output} / {outdir} を置換して実行する．\n"
        ),
    )
    ap.add_argument("input", metavar="INPUT.pptx", help="変換する pptx")
    ap.add_argument(
        "-o", "--output", metavar="PATH",
        help="出力する PDF（既定：入力と同じ場所・同じ basename の .pdf）")
    ap.add_argument(
        "--converter", metavar="NAME|COMMAND",
        help=f"変換器（auto / powerpoint / libreoffice / 任意のコマンド行）．"
             f"環境変数 {ENV_CONVERTER} を上書きする")
    ap.add_argument(
        "--timeout", metavar="SEC", type=float,
        help=f"変換を諦めるまでの秒数（0 で無制限）．環境変数 {ENV_TIMEOUT} を上書き"
             f"する．無指定なら，端末が tty のときは無制限，それ以外は 180 秒")
    ap.add_argument(
        "--unattended", action="store_true",
        help="端末が tty でも「人は見ていない」として扱う（cron・エディタのタスク用）．"
             "待ちを打ち切り，止まっても PowerPoint を前面に出さない")
    ap.add_argument("-q", "--quiet", action="store_true",
                    help="成功時に何も出力しない")
    ap.add_argument("--version", action="version", version=f"pptx2pdf {__version__}")
    return ap


def main(argv: list[str] | None = None) -> None:
    args = _build_parser().parse_args(argv)
    # 空の --output は「指定なし」と区別が付かないまま黙って別の場所へ書いてしまう
    # （`-o "$PDF_OUT"` で変数が未設定のときに起こる）．誤りとして弾く．
    if args.output is not None and not args.output.strip():
        raise SystemExit("pptx2pdf: --output requires a path")
    dst = args.output or default_pdf_path(args.input)
    converter = args.converter or os.environ.get(ENV_CONVERTER)
    try:
        convert(args.input, dst, converter, args.timeout,
                unattended=args.unattended)
    except PdfError as e:
        raise SystemExit(f"pptx2pdf: {e}")
    except KeyboardInterrupt:
        # 上限なしで待っているときの Ctrl-C は「止めたい」であって異常ではない．
        # トレースバックを出さずに，シェルの慣習どおりの終了コードで終わる．
        sys.stderr.write("\npptx2pdf: interrupted\n")
        raise SystemExit(130)
    if not args.quiet:
        print(f"saved: {dst}", flush=True)


if __name__ == "__main__":
    main()
