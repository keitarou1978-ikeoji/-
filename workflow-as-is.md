# 現状ワークフロー（As-Is）

## 全体フロー

```mermaid
flowchart TD
    subgraph 受注["📥 受注（販売先により異なる）"]
        FAX["FAX受注\n（紙で届く）"]
        EMAIL["メール受注\n（メール確認）"]
        SYSTEM["システム受注\n（取引先システム）"]
    end

    subgraph 受注処理["📝 受注処理（すべて手作業）"]
        PRINT["受注書を紙で打ち出し"]
        EXCEL["Excelの販売管理表に\n手入力"]
        CHECK["受注内容の確認・照合"]
    end

    subgraph 発注["📤 出荷・発注"]
        ORDER_PAPER["発注書を紙で作成"]
        FAX_SEND["出荷先へFAXで発注"]
        CONFIRM["出荷先からの確認待ち"]
    end

    subgraph 納品管理["📦 納品・請求"]
        DELIVERY["納品確認"]
        INVOICE["請求書作成"]
        PAYMENT["入金確認"]
    end

    FAX --> PRINT
    EMAIL --> PRINT
    SYSTEM --> PRINT

    PRINT --> EXCEL
    EXCEL --> CHECK
    CHECK --> ORDER_PAPER
    ORDER_PAPER --> FAX_SEND
    FAX_SEND --> CONFIRM

    CONFIRM --> DELIVERY
    DELIVERY --> INVOICE
    INVOICE --> PAYMENT

    style FAX fill:#ffcccc,stroke:#cc0000
    style PRINT fill:#ffcccc,stroke:#cc0000
    style EXCEL fill:#ffcccc,stroke:#cc0000
    style ORDER_PAPER fill:#ffcccc,stroke:#cc0000
    style FAX_SEND fill:#ffcccc,stroke:#cc0000
    style EMAIL fill:#fff3cd,stroke:#cc9900
    style SYSTEM fill:#d4edda,stroke:#28a745
```

### 凡例

- 🔴 赤：手作業・紙ベースの非効率なプロセス
- 🟡 黄：一部デジタルだが改善余地あり
- 🟢 緑：すでにデジタル化済み

---

## 販売先別の受注パターン

```mermaid
flowchart LR
    subgraph パターンA["パターンA：FAX受注"]
        A1["販売先がFAXで\n注文書を送信"] --> A2["担当者がFAXを\nピックアップ"]
        A2 --> A3["紙の注文書を見ながら\nExcelに手入力"]
        A3 --> A4["紙の発注書を作成\nFAXで出荷先へ送信"]
    end

    subgraph パターンB["パターンB：メール受注"]
        B1["販売先がメールで\n注文を送信"] --> B2["担当者がメールを確認\n注文内容を確認"]
        B2 --> B3["メール内容を見ながら\nExcelに手入力"]
        B3 --> B4["紙の発注書を作成\nFAXで出荷先へ送信"]
    end

    subgraph パターンC["パターンC：システム受注"]
        C1["販売先のシステムから\n注文データ連携"] --> C2["システムで\n受注データを確認"]
        C2 --> C3["受注データを\nExcelに転記"]
        C3 --> C4["紙の発注書を作成\nFAXで出荷先へ送信"]
    end

    style A1 fill:#ffcccc,stroke:#cc0000
    style A2 fill:#ffcccc,stroke:#cc0000
    style A3 fill:#ffcccc,stroke:#cc0000
    style A4 fill:#ffcccc,stroke:#cc0000
    style B1 fill:#fff3cd,stroke:#cc9900
    style B2 fill:#fff3cd,stroke:#cc9900
    style B3 fill:#ffcccc,stroke:#cc0000
    style B4 fill:#ffcccc,stroke:#cc0000
    style C1 fill:#d4edda,stroke:#28a745
    style C2 fill:#d4edda,stroke:#28a745
    style C3 fill:#ffcccc,stroke:#cc0000
    style C4 fill:#ffcccc,stroke:#cc0000
```

> **共通の課題**: どのパターンでも最終的に「Excelへの手入力」と「紙のFAX発注」に集約される

---

## 業務の時間配分（推定）

```mermaid
pie title 1件あたりの受注処理にかかる作業割合
    "FAX/メール確認・仕分け" : 10
    "受注書の打ち出し" : 5
    "Excelへの手入力" : 30
    "入力内容の照合・確認" : 15
    "発注書の作成" : 15
    "FAX送信・確認" : 10
    "ファイリング・保管" : 15
```

---

## 現状の課題まとめ

| # | 課題 | 影響 | 発生頻度 |
|---|------|------|----------|
| 1 | 受注チャネルがバラバラ（FAX/メール/システム） | 確認漏れ・対応遅延 | 毎日 |
| 2 | すべての受注をExcelに手入力 | 入力ミス・二重入力・時間浪費 | 毎日 |
| 3 | 発注がFAX（紙）ベース | 送信ミス・確認遅延・紙コスト | 毎日 |
| 4 | 受注～発注まで人手に依存 | 属人化・休暇時の業務停滞 | 常時 |
| 5 | データが分散（紙・Excel・メール） | 状況把握が困難・集計に時間 | 常時 |
