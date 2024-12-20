# Anything Markdown

## 概要

Anything Markdown は、さまざまなファイル形式をMarkdown形式に変換するためのGUIアプリケーションです。  
Pythonの`Tkinter`ライブラリを使用して構築されており、`MarkItDown`ライブラリを活用して高品質なMarkdown変換を実現します。  
さらに、OpenAIの大規模言語モデル（LLM）を利用して画像の説明を自動生成するオプションも提供しています。

## 主な機能

- **多様なファイル形式のサポート**: PDF、PowerPoint、Word、Excel、画像、HTML、CSV、JSON、XML、ZIP、テキストファイルなど、幅広い形式をMarkdownに変換可能。
- **LLMによる画像説明**: OpenAIのLLMを使用して、画像ファイルの説明を自動生成。
- **仮想環境での実行**: `.devcontainer`を使用して、一貫した開発環境を提供。

## 必要条件

- **Docker**: コンテナ化された開発環境を構築するために必要です。
- **Visual Studio Code (VS Code)**: Remote - Containers拡張機能を使用して開発環境にアクセス。
- **Docker Compose**: 一部の設定で必要になる場合があります。
- **OpenAI APIキー** (オプション): 画像の説明にLLMを使用する場合に必要です。

## インストールとセットアップ

### 1. リポジトリのクローン

```bash
git clone https://github.com/s-kuniyoshi/anything_markdown.git
cd markdown-converter
```

### 2. Dockerコンテナのビルドと起動
1. VS Codeを開く
2. Remote containers拡張機能のインストール
3. コンテナとしてプロジェクトを再オープン
    - Ctrl+Shift+P→コンテナで開くみたいなやつ
4. ビルドまち

### 3. 起動
1. terminalを開く
2. python main.pyを実行

## FAQ

### なんかGUIが違うけど？
しゃーない、Linuxの仮想環境で動いてるから。  
環境に依存したくないのだ

### ホストマシンのディレクトリにアクセスできないんだけど
.devcontainer/devcontainer.jsonの設定変えてくだされ  
特にWindows環境では、USERPROFILE 環境変数を使用してホストのホームディレクトリをマウントしてる
```
"mounts": [
    "source=${localEnv:USERPROFILE},target=/host,type=bind,consistency=cached"
]
```