# pptx2pdf

PowerPoint のファイル（pptx）を PDF に変換するコマンドです。
**実 PowerPoint**・LibreOffice・任意の変換ツールを使い分けられます。

依存は標準ライブラリだけです（python-pptx も要りません）。
pptx の中身は読まず、変換ツールを起こして結果を受け取ることに徹しています。

```bash
pptx2pdf slide.pptx                  # slide.pdf を隣に作る
pptx2pdf slide.pptx -o preview.pdf   # 出力先を指定する
pptx2pdf slide.pptx --converter libreoffice
pptx2pdf *.pptx                      # まとめて変換する
```

## インストール

```bash
# 推奨：pipx で隔離環境へインストール（PATH に pptx2pdf が入る）
pipx install git+https://github.com/toshi0806/pptx2pdf.git

# 手元のチェックアウトから
pipx install .
pip install .
```

インストールせずに試すなら `python3 -m pptx2pdf slide.pptx` でも動きます
（依存が無いので、リポジトリを取ってくるだけで実行できます）。

## コマンドラインリファレンス

| オプション | 説明 |
| --- | --- |
| `INPUT.pptx` | 変換する pptx。**複数指定できます** |
| `-o PATH` / `--output PATH` | 出力する PDF。既定は入力と同じ場所・同じ名前の `.pdf`。入力が1つのときだけ指定できます |
| `--converter NAME\|COMMAND` | 変換器。`auto`（既定）/ `powerpoint` / `libreoffice` / 任意のコマンド行。環境変数 `PPTX2PDF_CONVERTER` を上書き |
| `--timeout SEC` | 変換を諦めるまでの秒数（`0` で無制限）。環境変数 `PPTX2PDF_TIMEOUT` を上書き |
| `--unattended` | 端末が tty でも「人は見ていない」として扱う（cron やエディタのタスク用） |
| `-q` / `--quiet` | 成功時に何も出力しない |
| `--version` | バージョンを表示する |

- 出力先には**存在するディレクトリ**を指定してください（無い場所を指すと失敗します。
  打ち間違いで勝手にディレクトリを作ることはしません）。
- **PDF は変換中も一瞬たりとも消えません。** 作業用のディレクトリで変換してから
  `os.replace` で置き換えます。多くの PDF ビューアはファイルが消えた時点で監視をやめて
  しまうので、「消してから書く」と開き直すまで自動リロードされなくなるためです。
- **失敗したときは出力 PDF を消します。** 古い内容を新しい出力と取り違えないためです。
  終了コードは1で、理由は `pptx2pdf: …` の1行で stderr に出ます。
- **入力を複数渡したときは、1つ失敗しても残りを変換します。** 失敗したものは名前と理由を
  stderr に出し、最後に `pptx2pdf: 1 of 3 failed` と伝えて終了コード1で終わります
  （`-q` を付けても失敗の報告は出ます）。

## 変換器

既定の `auto` は**実 PowerPoint → LibreOffice** の順に、使えるものを選びます。
**選んだ変換器が失敗しても次へは落としません。** 忠実度の違う PDF が黙って出てくるのは
避けたいのと、止まる原因（オートメーションの承認、ライセンス未認証など）は多くの場合
利用者が直せるものだからです。

| 指定 | 使うもの |
| --- | --- |
| `auto` | 実 PowerPoint（macOS: AppleScript / Windows: COM）、無ければ LibreOffice |
| `powerpoint` | 実 PowerPoint のみ |
| `libreoffice` | `soffice --headless` のみ |
| コマンド行 | 指定したツール |

**忠実度は変換器で決まります。** LibreOffice の出力はテーマフォントの解決差などで
実 PowerPoint と一致しません（当たり確認には使えますが、最終確認の代わりにはなりません）。

**任意のコマンド**も指定できます。`{input}`（pptx）/ `{output}`（PDF のパス）/ `{outdir}`
（その親ディレクトリ）が実際のパスに置き換わります。出力先を指定できないツールなら、
プレースホルダは書かなくて構いません。pptx のパスが末尾に付いて実行されるので、
**pptx と同じディレクトリに同じ名前で `.pdf` を書く**ツール（`slide.pptx` なら
`slide.pdf`）なら、そのまま使えます。

```bash
export PPTX2PDF_CONVERTER='soffice --headless --convert-to pdf --outdir {outdir} {input}'
```

## 変換が終わらないときの待ち方

**端末から実行しているかどうかで変わります。** 端末から実行しているなら待ち続けます
（承認ダイアログのように、応答すれば進むことがあるためです。やめたいときは `Ctrl-C`）。
cron やエディタから実行したときのように**誰も応答できない場合は180秒で諦めます**。
`--timeout SEC` と `PPTX2PDF_TIMEOUT` で変えられます（`0` で無制限）。

端末から実行していても人が見ていないとき（エディタのタスク、監視ループ）は
`--unattended` を付けてください。上限が効くようになり、止まっても PowerPoint を
前面に出しません。

## macOS で PowerPoint に変換させるとき

- **初回に「オートメーション」の承認**が要ります。
  PowerPoint を操作してよいか尋ねるダイアログが出るので、一度許可してください。
  承認は**呼び出し元アプリごと**（Terminal / iTerm / VS Code など）に別管理なので、
  実行元を変えると再度承認が要ります。
- **変換中に PowerPoint は表示しません。** そのため、承認ダイアログのような応答待ちは
  無音の停止にしか見えません。**30秒たっても終わらなければ、その旨を stderr に出します**
  （端末から実行しているときは PowerPoint も前面に出します）。
- **すでに PowerPoint を開いているときは、変換中にウィンドウが出ます。** 作業中の
  PowerPoint を勝手に隠さないための割り切りです（フォーカスは奪いません）。
- **壊れた pptx を渡すと、その後の変換まで失敗することがあります。** PowerPoint が
  開けなかったファイルのダイアログを抱えたままになるためで、隠して動かしている以上
  それは画面に見えません。以後 `powerpoint` 経路が失敗し続けるときは、PowerPoint を
  前面に出してダイアログに答えるか、いったん終了させてください
  （`pptx2pdf *.pptx` のようにまとめて変換するときに出会いやすい挙動です）。
- 変換は PowerPoint 自身のサンドボックスコンテナの中で行うので、入出力がどこにあっても
  ファイルアクセスの承認ダイアログは出ません。

## ライブラリとして使う

```python
import pptx2pdf

pptx2pdf.convert("deck.pptx", "deck.pdf", None)         # None / "auto" で自動選択
pptx2pdf.convert("deck.pptx", "deck.pdf", "libreoffice", timeout=60)
```

失敗は `pptx2pdf.PdfError` です。出力先を省略したいときは
`pptx2pdf.default_pdf_path("deck.pptx")` が既定のパスを返します。

別のコマンドの一部として組み込むときは、利用者に見える名前を差し替えられます。
警告の接頭辞や「上限の延ばし方」の案内が、利用者が実際に打ったコマンドのものになります。

```python
pptx2pdf.set_program_name("md2pptx")
pptx2pdf.set_hints(timeout="--pdf-timeout / MD2PPTX_PDF_TIMEOUT",
                   converter="--pdf-converter")
```

## 由来

もとは [md2pptx](https://github.com/toshi0806/md2pptx) の `--pdf` の実装
（`md2pptx/pdf.py` と `md2pptx/workdir.py`）です。python-pptx に依存していなかったので、
単体のコマンドとして切り出しました。md2pptx はこのパッケージに依存しています。
**コード中の Issue 番号は md2pptx リポジトリのもの**で、設計の経緯はそちらに残っています。

## ライセンス

MIT
