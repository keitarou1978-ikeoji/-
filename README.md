# 販売業務ワークフロー可視化プロジェクト

販売業務の現状ワークフローを可視化し、標準化・自動化・DX化の改善ポイントを明確にするためのドキュメントです。

## ファイル構成

| ファイル | 内容 |
|---------|------|
| [workflow-as-is.md](./workflow-as-is.md) | 現状ワークフロー（As-Is） |
| [workflow-to-be.md](./workflow-to-be.md) | 改善後ワークフロー（To-Be） |
| [improvement-plan.md](./improvement-plan.md) | 改善ポイント一覧と優先度マトリクス |

## 閲覧方法

Mermaid記法を使用しています。以下の方法で図を確認できます：

- **GitHub** 上で直接表示（Mermaidに対応済み）
- **VS Code** で [Mermaid Preview](https://marketplace.visualstudio.com/items?itemName=bierner.markdown-mermaid) 拡張を使用
- **[Mermaid Live Editor](https://mermaid.live/)** にコードを貼り付け

---

## 個人資金管理ツール（人生奪還PJ）

Excel + VBA による個人向け資金管理ツール。設計仕様と Phase 1 実装を同梱しています。

| ファイル | 内容 |
|---------|------|
| [personal_finance_tool_prompt_v1.md](./personal_finance_tool_prompt_v1.md) | 設計仕様書（全10シート・全8フェーズ） |
| `人生奪還PJ_資金管理ツール_v1.xlsm` | 成果物ブック（Phase 1: 設定／手入力シート） |
| [build/build_workbook.py](./build/build_workbook.py) | ブック生成スクリプト |
| [build/verify_sampledata.py](./build/verify_sampledata.py) | テストデータ生成ロジックの検証 |
| [vba/](./vba/) | VBAモジュール（modConfig / modManualEntry / modUtils / shtManual） |
| [docs/SETUP_VBA.md](./docs/SETUP_VBA.md) | VBAインポート・ボタン割当の手順 |
