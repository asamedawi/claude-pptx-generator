# claude-pptx-generator

Claude Code を使って PowerPoint 資料を生成するためのプロジェクトです。

「○○ についての資料を作成してほしい」と依頼すると、AI がテーマを読み取り、目的・対象者・所要時間を確認したうえで骨子を作成し、合意後に `output/output.pptx` を生成します。

## 概要

```text
claude-pptx-generator/
|-- output/
|   |-- slides.md     生成前のスライド原稿
|   `-- output.pptx   生成される PowerPoint ファイル
|-- templates/
|   `-- template.pptx 使用する PowerPoint テンプレート
|-- CLAUDE.md         AI の動作ルールと依頼テンプレート
|-- generate_pptx.py  Markdown から PowerPoint を生成するスクリプト
|-- README.md         このプロジェクトの使い方
|-- references.md     事前に登録しておく参考文献一覧
`-- requirements.txt  Python 依存関係
```

## セットアップ

- `pip install -r requirements.txt` で依存関係をインストールします。
- `templates/template.pptx` が存在しない場合は配置します。

## 参考文献の設定方法

参考文献は対話のたびに入力するのではなく、`references.md` にあらかじめ記載しておきます。AI は資料作成時にこのファイルを最初に参照します。

### 記載場所

- ファイル: `references.md`
- `記載例：`

```md
# 参考文献

- AWS Amplify Gen 2 ドキュメント: https://docs.amplify.aws/
- Amazon DynamoDB ドキュメント: https://docs.aws.amazon.com/dynamodb/
```

### 運用ルール

- 定常的に使う出典は `references.md` に追記します。
- 資料ごとに追加したい参考文献がある場合は、依頼文の中で別途指定できます。

## 使い方

1. このプロジェクトを開きます。
2. 必要に応じて `references.md` を更新します。
3. Claude Code に「○○ についての資料を作成してほしい」と依頼します。
4. AI の質問に対して、目的・対象者・所要時間を回答します。
5. 骨子を確認し、プランモードで合意します。
6. AI がスライド内容を整理し、`generate_pptx.py` を実行して `output/output.pptx` を生成します。
7. 生成後、PowerPoint が自動で開きます。

## 補足

AI の詳細な動作ルールは `CLAUDE.md` に定義されています。
