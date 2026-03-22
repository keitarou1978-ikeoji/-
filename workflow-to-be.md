# 改善後ワークフロー（To-Be）

## Phase 1：受注の一元化（短期 〜3ヶ月）

```mermaid
flowchart TD
    subgraph 受注["📥 受注チャネル"]
        FAX["FAX受注"]
        EMAIL["メール受注"]
        SYSTEM["システム受注"]
    end

    subgraph 統合["🔄 受注の一元化"]
        FAX_OCR["FAX → 複合機でPDF化\n→ OCRで自動読取"]
        EMAIL_PARSE["メール → 受注管理システムへ\n自動取り込み"]
        API["システム → API連携で\n自動取り込み"]
        DB["受注管理システム\n（一元管理DB）"]
    end

    FAX --> FAX_OCR
    EMAIL --> EMAIL_PARSE
    SYSTEM --> API

    FAX_OCR --> DB
    EMAIL_PARSE --> DB
    API --> DB

    style FAX_OCR fill:#d4edda,stroke:#28a745
    style EMAIL_PARSE fill:#d4edda,stroke:#28a745
    style API fill:#d4edda,stroke:#28a745
    style DB fill:#cce5ff,stroke:#0066cc
```

> **効果**: Excel手入力をなくし、受注データを一箇所に集約

---

## Phase 2：発注の自動化（中期 3〜6ヶ月）

```mermaid
flowchart TD
    DB["受注管理システム"] --> AUTO_ORDER["受注データから\n自動で発注書を生成"]
    AUTO_ORDER --> ROUTE{"出荷先の\n対応状況"}

    ROUTE -->|"システム対応済み"| EDI["EDI/API連携で\n自動発注"]
    ROUTE -->|"メール対応可"| AUTO_MAIL["発注書をメールで\n自動送信"]
    ROUTE -->|"FAXのみ対応"| AUTO_FAX["クラウドFAXで\n自動送信"]

    EDI --> STATUS["ステータス自動更新"]
    AUTO_MAIL --> STATUS
    AUTO_FAX --> STATUS

    STATUS --> DASHBOARD["ダッシュボードで\nリアルタイム進捗確認"]

    style AUTO_ORDER fill:#d4edda,stroke:#28a745
    style EDI fill:#d4edda,stroke:#28a745
    style AUTO_MAIL fill:#d4edda,stroke:#28a745
    style AUTO_FAX fill:#fff3cd,stroke:#cc9900
    style DASHBOARD fill:#cce5ff,stroke:#0066cc
```

> **効果**: 手作業の発注書作成・FAX送信を廃止し、自動化

---

## Phase 3：全体最適化（長期 6〜12ヶ月）

```mermaid
flowchart TD
    subgraph 受注["📥 受注（自動取り込み）"]
        CHANNELS["全チャネル\n自動取り込み済み"]
    end

    subgraph コア["⚙️ 販売管理システム（中核）"]
        ORDER_MGT["受注管理"]
        STOCK["在庫管理"]
        SHIPPING["出荷管理"]
        INVOICE_MGT["請求管理"]
    end

    subgraph 出力["📊 アウトプット"]
        DASHBOARD["リアルタイム\nダッシュボード"]
        ALERT["異常値アラート\n（売上・在庫）"]
        REPORT["自動レポート\n（日次/月次）"]
    end

    CHANNELS --> ORDER_MGT
    ORDER_MGT --> STOCK
    ORDER_MGT --> SHIPPING
    SHIPPING --> INVOICE_MGT

    ORDER_MGT --> DASHBOARD
    STOCK --> ALERT
    INVOICE_MGT --> REPORT

    style CHANNELS fill:#d4edda,stroke:#28a745
    style ORDER_MGT fill:#cce5ff,stroke:#0066cc
    style STOCK fill:#cce5ff,stroke:#0066cc
    style SHIPPING fill:#cce5ff,stroke:#0066cc
    style INVOICE_MGT fill:#cce5ff,stroke:#0066cc
    style DASHBOARD fill:#e8daef,stroke:#8e44ad
    style ALERT fill:#e8daef,stroke:#8e44ad
    style REPORT fill:#e8daef,stroke:#8e44ad
```

---

## As-Is → To-Be 比較

```mermaid
flowchart LR
    subgraph 現状["❌ 現状（As-Is）"]
        direction TB
        NOW1["FAX/メール/システム\nバラバラに受注"]
        NOW2["紙に打ち出し"]
        NOW3["Excelに手入力"]
        NOW4["紙の発注書作成"]
        NOW5["FAXで発注"]
        NOW1 --> NOW2 --> NOW3 --> NOW4 --> NOW5
    end

    subgraph 改善後["✅ 改善後（To-Be）"]
        direction TB
        NEW1["全チャネル\n自動取り込み"]
        NEW2["受注管理システムに\n自動登録"]
        NEW3["発注書を\n自動生成・送信"]
        NEW4["ダッシュボードで\nリアルタイム確認"]
        NEW1 --> NEW2 --> NEW3 --> NEW4
    end

    現状 -- "DX化" --> 改善後

    style NOW1 fill:#ffcccc,stroke:#cc0000
    style NOW2 fill:#ffcccc,stroke:#cc0000
    style NOW3 fill:#ffcccc,stroke:#cc0000
    style NOW4 fill:#ffcccc,stroke:#cc0000
    style NOW5 fill:#ffcccc,stroke:#cc0000
    style NEW1 fill:#d4edda,stroke:#28a745
    style NEW2 fill:#d4edda,stroke:#28a745
    style NEW3 fill:#d4edda,stroke:#28a745
    style NEW4 fill:#d4edda,stroke:#28a745
```
