# TransCYPlate

**作者**: 紫波レント / Roent Shiba  
**ライセンス**: CC BY-NC-SA 4.0 （帰属: 紫波レント / Roent Shiba）  
（詳しくは下部の「ライセンス」参照）

## 概要

TransCYPlate は、ドイツ語テキストをリアルタイムに英語（EN）と日本語（JA）へ逐次翻訳し、出現した単語を Word-Flash で学習できる軽量デスクトップツールです。翻訳は LM Studio クライアント（ローカル/ネットワーク経由）に接続して実行します。スペイン語（ES）・フランス語（FR）は軽量化のためデフォルトで無効化されています。

## 主な機能

* German → English → Japanese（順序固定）の逐次翻訳キュー（文単位）
* 出現単語を抽出して個別に訳・補強する Word-Flash（単語キュー）
* Word CSV 管理（`log/word.csv`）: `word,en,ja,count,skip`（skip=1 の単語は自動表示をスキップ）
* Re-display（手動再表示）機能：skip フラグに関係なく再表示可能。skip==1 の語を再表示した際は `★` を登場回数の前に表示
* Skip Toggle：選択単語の skip フラグを 0↔1 で切替
* WordFlash の表示を PNG 保存（`log/FlashPNG/yyyyMMdd/(word).png`）

  * 新規単語（登場回数が 1 のとき）の自動 PNG 保存機能あり
* GUI に各キュー（Sentence/Word）サイズと現在処理中の項目を表示
* ウィンドウ更新時にメインウィンドウのフォーカスを奪わない（利便性の調整）
* prompt の改行を除去するサニタイズ（改行による表示崩れを防止）

## 依存関係

* Python 3.8+
* lmstudio ライブラリ（LM Studio を使う場合）
* Pillow（PIL） — WordFlash を PNG で保存する場合に推奨
* Tkinter（標準 GUI）

> WSL 環境やリモートディスプレイでは `ImageGrab` が動作しない場合があります。PNG 保存を使う場合はネイティブのデスクトップ環境（Windows / X サーバ）を推奨します。

## インストール（例）

```bash
# 必要なパッケージをインストール
pip install lmstudio pillow
```

## 実行方法

LM StudioのDeveloperでgamma-3n-e4bかgpt-oss-20bのサーバーを起動する

```bash
python3 TransCYPlate_gemma3n_v1_0_18.py
```
もしくは
```bash
python3 TransCYPlate_GPToss_v1_0_18.py
```


## 設定

* 実行フォルダに `simple_live_translator.config.json` を作成するか、GUI の Connection セクションで設定を編集して Save を押してください。主な設定項目:

  * `SERVER_API_HOST` — LM Studio サーバ（例: `localhost:1234`）
  * `MODEL_NAME` — 使用するモデル名（例: `google/gemma-3n-e4b`）
  * `TEMPERATURE` — 翻訳の温度
  * `MAX_TOKENS` — 最大トークン（0 は自動）

デフォルトでは `DISABLED_LANGS = {"es","fr"}` によりスペイン語・フランス語は無効です。必要ならソース内でこの設定を変更できます。

## `log/word.csv` の仕様

CSV ヘッダ（最新版）:

```
word,en,ja,count,skip
```

* `word` — 単語
* `en` — English 候補（セミコロン区切り）
* `ja` — Japanese 候補（セミコロン区切り）
* `count` — 出現数（自動でインクリメント）
* `skip` — 0/1 フラグ。1 の場合は自動 WordFlash をスキップ（ただし手動 Re-display は可能）

古い 4 列や 2 列形式にも互換的に対応します。

## GUI の簡単な説明

* Main Window: 接続設定、入力ボックス（ドイツ語）、Queue 状態、Saved Words コンボボックス（`word (count)` 表示）、Re-display / Skip / Save PNG ボタン、Q\&A 欄など。
* Display Window: 翻訳された文を表示。左上に言語ステータス（英/日/西/仏）と右上に `S 数字  W 数字`（Sentence/Word キュー数）を白字で表示。
* Word Flash Window: 選択単語のカウント、ドイツ語・英語・日本語を大きく表示。`★` は **手動で Re-display** したときに、その単語が `skip==1` の場合に登場回数の前に表示されます。

## 保存（PNG）

* 手動で `Save PNG` を押すと WordFlash をキャプチャし、`log/FlashPNG/YYYYMMDD/(sanitized_word).png` に保存します。
* 新しい単語が最初に表示された（count == 1）場合、自動で同ディレクトリに PNG を保存します（環境に依存するため動かない場合あり）。

## 動作ポリシー / 注意

* このツールはローカルで LM Studio の LLM を呼び出す前提です。インターネット上の API キーを直接埋め込む実装はしていません。
* `ImageGrab` を使った画面キャプチャはプラットフォーム依存です。WSL + X 環境では失敗することがあるため、その場合はネイティブ環境か代替手段を検討してください。

## トラブルシューティング

* `Pillow`（ImageGrab）が無い、または PNG 保存でエラーが出る → `pip install pillow` を行う。WSL の場合は画面キャプチャ不可の可能性あり。
* LM Studio に接続できない → `simple_live_translator.config.json` の `SERVER_API_HOST` を確認、LM Studio サーバが起動しているか確認。

## 貢献

バグ報告や改善提案は Issue / PR を歓迎します。主に以下の点が今後の候補です:

* ES/FR の再有効化と並行ワーカーの最適化
* WordFlash を複数デザインで保存する機能
* headless 環境での PNG 出力対応強化

## 謝辞

ChatGPT 5 Thinking 様
本プロジェクトは、紫波レントがノーコードで制作し、コードの自動生成・改善提案に ChatGPT 5 Thinking を活用しました。

## ライセンス

このプロジェクトは **CC BY-NC-SA 4.0** の下で公開します。
帰属: 紫波レント / Roent Shiba

**短い表示例（LICENSE ファイルに追記してください）**:

```
Copyright (c) Roent Shiba (紫波レント)

This work is licensed under the Creative Commons Attribution-NonCommercial-ShareAlike 4.0 International License.
To view a copy of this license, visit http://creativecommons.org/licenses/by-nc-sa/4.0/
```

---
