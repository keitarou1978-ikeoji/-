# 販売業務ワークフロー可視化プロジェクト

販売業務の現状ワークフローを可視化し、標準化・自動化・DX化の改善ポイントを明確にするためのドキュメントです。

## ファイル構成

| ファイル | 内容 |
|---------|------|
| [workflow-as-is.md](./workflow-as-is.md) | 現状ワークフロー（As-Is） |
| [workflow-to-be.md](./workflow-to-be.md) | 改善後ワークフロー（To-Be） |
| [improvement-plan.md](./improvement-plan.md) | 改善ポイント一覧と優先度マトリクス |
| [claude-qa.md](./claude-qa.md) | Claudeへの質問と回答のログ |

## ふるさと納税 商品登録テンプレート

新しい返礼品を登録する際は以下の手順で提案ファイルを作成してください。

```bash
# 新しい商品の提案ファイルを生成
bash templates/new-proposal.sh "商品名"
# → proposals/YYYYMMDD_商品名.md が作成されます
```

- **テンプレート**: [`templates/satofuru-proposal-template.md`](./templates/satofuru-proposal-template.md)
- **提案ファイル保存先**: `proposals/` フォルダ

---

## 閲覧方法

Mermaid記法を使用しています。以下の方法で図を確認できます：

- **GitHub** 上で直接表示（Mermaidに対応済み）
- **VS Code** で [Mermaid Preview](https://marketplace.visualstudio.com/items?itemName=bierner.markdown-mermaid) 拡張を使用
- **[Mermaid Live Editor](https://mermaid.live/)** にコードを貼り付け
