# tumitate（積立・乗り換え判定ツール）

## 目的
保険を解約して別の商品（預金・国債・NISAなど）に積み立て直すべきかを、
利回り・税引後の受取額で比べて判断するためのツール。

## 構成
- `app/index.html` … 比較アプリ本体（1ファイル完結。ブラウザで開くだけで動く）
  - どの保険会社の試算書でも入力できる。複数の保険を保存して切り替え可能
  - 保険の初期データは `/* POLICY_DATA_START */ 〜 /* POLICY_DATA_END */` の `DEFAULT_POLICIES`
  - **公開リポジトリなので、ここには架空のサンプルだけを入れる**
  - 入力した値はブラウザ（localStorage、キー `hoken-norikae-v3`）にだけ保存される
  - 初期データを変えたら `DATA_VERSION` を上げる（利用者が編集していない初期の保険だけ入れ替わる）
- `excel/make_xlsx.py` … シンプルな Excel 版（A：今の保険、B：円建て一時払、C：ドル建て一時払 の返戻率・年利回りを比べる）を作るスクリプト
  - 引数なし → 架空サンプル入りのひな形 `excel/hoken-rimawari-hikaku-sample.xlsx`
  - `python excel/make_xlsx.py 本人データ.json 出力.xlsx` → 本人用（JSONも出力もリポジトリに置かない）
  - 作ったら xlsx スキルの `recalc.py` で計算と式のエラーを確認する
- `PROGRESS.md` … 作業の進み具合
- `KNOWLEDGE.md` … 計算ルールや金利などの学び

## 公開しているページ（claude.ai の Artifact）
- 「保険の乗り換え判定」自分用（非公開）：`app/index.html` の POLICY_DATA を本人データに差し替え、
  `<!--SKELETON-->` の行を取り除いて公開する。本人データはリポジトリに置かない
- 「保険の乗り換え判定」共有用：`app/index.html` から `<!--SKELETON-->` の行を取り除いてそのまま公開
- 更新するときは、Artifact の一覧から URL を探し、読み込んでから同じ URL に公開する
- どちらも公開時に `capabilities: {downloads: true, db: {}, user: {}, sample: {images: true}}` を付ける
  - downloads：比較レポートのPDFとデータファイルの保存 / sample（画像あり）：写真からClaudeで表を読み取る
  - db・user：スマホ→パソコンの受け渡し（本人専用の保存場所 `data/users/<id>/scan-inbox`。パソコンが合言葉つきで「待っている」と書き、受け取ったら消す）
- QRコードに入れるURLは POLICY_DATA ブロックの `APP_URL`。共有用は共有ページのURL、自分用は自分用ページのURLに差し替える

## 自分のサイト（会員サイト）に置く版
- `app/index.html` の `/* SERVER_CONFIG_START */` の `SERVER_API` に、そのサイトの窓口のパスを入れたコピーを置く（リポジトリでは空のまま）
- `SERVER_API` があると、claude.ai の機能の代わりにサーバーの窓口を使う
  - 受け渡し：`{SERVER_API}/handoff/new・status・put・done・get・end`（QRは `?h=番号`）
  - 読み取り：`{SERVER_API}/surrender` に `{files, tiles}` を送り、OCR_PROMPT の形の JSON を受け取る
- 会員サイトのアドレスや設定は、公開リポジトリに書かない。置き方はユーザーに渡した引き継ぎ書で伝える

## ルール
- ユーザー本人の保険の数字は、このリポジトリにコミットしない
  （本人の数字入りの版は claude.ai の非公開 Artifact「保険の乗り換え判定」で管理）
- 金利など「仮の前提」は、画面の「前提と注意」に必ず書く
- 回答は短く、IT用語はかみ砕く。図にできるものは図にする
