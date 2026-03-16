# claude-pptx-generator

Claude Code を使って PowerPoint を生成する。

「○○について資料を作成して」と伝えると、AI が目的・対象者・所要時間などを質問する。その後、骨子を確認したうえでスライドを作成し、`output/output.pptx` を出力する。

## 概要

- **依頼例**: 「AI活用についての資料を作成してほしい」
- **AI の流れ**: テーマ読み取り → 質問（目的・対象者・所要時間・参考文献） → 骨子作成 → プラン確認 → スライド内容作成 → Markdown 出力 → `generate_pptx.py` で pptx 生成
- **テンプレート**: `templates/template.pptx`
- **出力先**: `output/output.pptx`

## セットアップ

- `pip install -r requirements.txt` で依存関係をインストール
- `templates/template.pptx` を配置する（プロジェクトに含まれていない場合は用意すること）

## 使い方

1. 任意のエディタでこのプロジェクトを開く
2. Claude Code に「○○についての資料を作成してほしい」と依頼する
3. AI の質問（目的・対象者・所要時間・参考文献など）に答える
4. 骨子を確認し、プランモードで合意する
5. AI がスライドを作成し、`generate_pptx.py` を実行して `output/output.pptx` を生成する
6. 生成完了後、PowerPoint が自動で開く

CLAUDE.md に AI 向けの詳細ルールと依頼テンプレートが定義されている。