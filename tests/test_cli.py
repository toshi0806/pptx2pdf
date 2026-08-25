#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""コマンドとしての入口を固定するテスト．

変換そのものは他のテストが見ているので、ここで守るのは **引数と環境変数をどう解決して
``convert`` へ渡すか**と、**失敗をどう終わらせるか**だけ．``convert`` を差し替えるので
外部プロセスは起こさない．
"""
from __future__ import annotations

import pytest

from pptx2pdf import cli, converter


@pytest.fixture
def calls(monkeypatch):
    """``cli`` が呼んだ ``convert`` の引数を記録する（変換はしない）．"""
    seen: list[dict] = []

    def fake(src, dst, conv, timeout=None, *, unattended=False):
        seen.append({"src": src, "dst": dst, "converter": conv,
                     "timeout": timeout, "unattended": unattended})

    monkeypatch.setattr(cli, "convert", fake)
    return seen


def test_the_pdf_lands_next_to_the_pptx_by_default(calls, tmp_path, capsys):
    """出力先を省いたら入力と同じ場所・同じ basename の .pdf．"""
    src = tmp_path / "deck.pptx"
    cli.main([str(src)])
    assert calls[0]["dst"] == str(tmp_path / "deck.pdf")
    assert f"saved: {tmp_path / 'deck.pdf'}" in capsys.readouterr().out


def test_an_empty_output_path_is_refused(calls, tmp_path):
    """空の ``-o`` は弾く．

    ``-o "$PDF_OUT"`` で変数が未設定のときに起こる．黙って既定の場所へ書くと、
    利用者は指定した場所を探し続けることになる．
    """
    with pytest.raises(SystemExit, match="--output requires a path"):
        cli.main([str(tmp_path / "deck.pptx"), "-o", "  "])
    assert calls == []


def test_the_environment_supplies_the_converter(calls, tmp_path, monkeypatch):
    """``--converter`` 無指定なら環境変数を使い、指定があればそちらが勝つ．"""
    monkeypatch.setenv(converter.ENV_CONVERTER, "libreoffice")
    cli.main([str(tmp_path / "deck.pptx")])
    assert calls[-1]["converter"] == "libreoffice"
    cli.main([str(tmp_path / "deck.pptx"), "--converter", "powerpoint"])
    assert calls[-1]["converter"] == "powerpoint"


def test_options_reach_the_conversion(calls, tmp_path):
    """``--timeout`` / ``--unattended`` はそのまま ``convert`` へ渡す．"""
    cli.main([str(tmp_path / "deck.pptx"), "--timeout", "7", "--unattended"])
    assert calls[0]["timeout"] == 7.0
    assert calls[0]["unattended"] is True


def test_a_failed_conversion_exits_with_the_reason(monkeypatch, tmp_path):
    """変換の失敗は ``pptx2pdf: <理由>`` の 1 行で終了コード 1．

    md2pptx の ``--pdf``（pptx は保存できているので警告で済ませる）と違い、ここでは
    PDF が唯一の成果物なので**失敗として終わる**．
    """
    def boom(*args, **kwargs):
        raise converter.PdfError("libreoffice failed: boom")

    monkeypatch.setattr(cli, "convert", boom)
    with pytest.raises(SystemExit) as e:
        cli.main([str(tmp_path / "deck.pptx")])
    assert str(e.value) == "pptx2pdf: libreoffice failed: boom"


def test_quiet_says_nothing_on_success(calls, tmp_path, capsys):
    """``-q`` は成功時に何も出さない（スクリプトから使うため）．"""
    cli.main([str(tmp_path / "deck.pptx"), "-q"])
    assert capsys.readouterr().out == ""
