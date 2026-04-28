# Dify API × Python 設計書

## 何を作るか
PythonからDify APIを叩いて、返答を受け取るスクリプト

## 処理の流れ
1. Pythonからメッセージを送る
2. Dify APIが受け取る
3. AIが返答を生成する
4. Pythonが返答を受け取って表示する

## 使うもの
- Python
- requests（HTTPリクエストを送るライブラリ）
- Dify API Key

## ファイル構成
Python/
└── dify_test/
    ├── main.py     ← 実行ファイル
    └── config.py   ← APIキーなどの設定

## 次のステップ
- [ ] Dify APIキーを確認する
- [ ] main.pyを書く
- [ ] 動かしてみる