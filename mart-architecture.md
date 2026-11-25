# Ticket Mart Architecture (工單 Mart 架構)

`mart_build.py` 把爬蟲寫入的年度 .db（唯讀）和人效項目清冊快照，全量重建成分層的 `mart/mart.db`。
本文件用四張圖說明：資料怎麼分層流動、表的分工、表與表怎麼關聯、一次建置怎麼執行。
文字說明與實測數字見 [README〈建置 Ticket Mart〉](../README.md)。

---

## 0. Color Legend (顏色圖例)

所有圖都使用同一組顏色。**每一層一種顏色**；紅、綠、黃只用在有語意的地方（失敗、成功、需要留意），不代表任何一層。

```mermaid
flowchart LR
    SRC["來源系統"]:::src
    RAW["L0 原始層<br/>不修改"]:::raw
    STG["L1 標準化層<br/>只整理、不加規則"]:::stg
    DIM["L2 維度表 dim_<br/>描述「誰、什麼、哪裡、何時」"]:::dim
    FACT["L2 事實表 fact_<br/>記錄「發生了什麼、多少」"]:::fact
    MART["L3 輸出層 mart_<br/>給特定使用者的形狀"]:::mart
    META["建置紀錄<br/>參數、log、驗收"]:::meta
    USER["使用者"]:::user
    DQ["需要留意<br/>資料品質待補"]:::warn
    OK["成功"]:::ok
    ERR["失敗"]:::err

    SRC ~~~ RAW ~~~ STG ~~~ DIM ~~~ FACT ~~~ MART
    META ~~~ USER ~~~ DQ ~~~ OK ~~~ ERR

    classDef src  fill:#FAFAFA,stroke:#9E9E9E,color:#212121,stroke-dasharray:4 3
    classDef raw  fill:#ECEFF1,stroke:#546E7A,color:#263238
    classDef stg  fill:#E0F2F1,stroke:#00897B,color:#004D40
    classDef dim  fill:#EDE7F6,stroke:#5E35B1,color:#311B92
    classDef fact fill:#E3F2FD,stroke:#1565C0,color:#0D47A1,stroke-width:2px
    classDef mart fill:#EFEBE9,stroke:#6D4C41,color:#3E2723,stroke-width:2px
    classDef meta fill:#F5F5F5,stroke:#BDBDBD,color:#616161,stroke-dasharray:2 2
    classDef user fill:#FFFFFF,stroke:#424242,color:#212121
    classDef warn fill:#FFF8E1,stroke:#F9A825,color:#5D4037
    classDef ok   fill:#E8F5E9,stroke:#2E7D32,color:#1B5E20
    classDef err  fill:#FFEBEE,stroke:#C62828,color:#B71C1C
```

| 顏色 | 代表 | 第一次看可以這樣理解 |
|---|---|---|
| 灰（虛線框） | 來源系統 | 資料從哪裡來，不在 mart 裡 |
| 藍灰 | L0 原始層 | 爬蟲抓回來的原樣資料，mart 只讀不改 |
| 青綠 | L1 標準化層 | 把欄位整理乾淨，不加任何判斷 |
| 紫 | L2 維度表 `dim_` | 用來「切」資料的角度：服務項目、經手人、群組、申請人、日期 |
| 藍（粗框） | L2 事實表 `fact_` | 被「量」的東西：工單、對話、經手紀錄，以及時數、點數 |
| 棕（粗框） | L3 輸出層 `mart_` | 已經整理成某個使用者要的樣子，可以直接拿去用 |
| 淺灰（點線框） | 建置紀錄 | 這次用了哪些參數、每張表幾列、驗收結果 |
| 黃 | 需要留意 | 對不到清冊的值，請清冊維護者補 |
| 綠 / 紅 | 成功 / 失敗 | 只出現在建置步驟圖 |

---

## 1. Data Flow (分層資料流)

由左往右讀：資料每往右一層就更「好用」。每一層只讀前一層；業務規則只出現在 L2（穩定的規則）與 L3（會變的政策）。

```mermaid
flowchart LR
    subgraph SRC["來源"]
        FS["Freshservice<br/>API v2 + 後台 view"]:::src
        XLSX["人效項目清冊.xlsx<br/>Fresh Item / Fresh Agent"]:::src
    end

    subgraph L0["L0 原始層（不修改）"]
        RAW[("db/vig_freshservice_YYYY.db<br/>ticket_conversations")]:::raw
        REFX["ref/effort_items.xlsx<br/>清冊快照"]:::raw
    end

    subgraph MART["mart/mart.db（每次全量重建）"]
        direction LR
        REF["ref_fresh_item<br/>ref_fresh_agent<br/>ref_meta（sha256）"]:::raw
        subgraph L1["L1 標準化｜10_stg.sql"]
            STG["stg_ticket<br/>stg_conversation<br/>stg_fresh_item<br/>stg_fresh_agent"]:::stg
        end
        subgraph L2["L2 核心｜20_core_dims.sql、21_core_facts.sql"]
            DIM["dim_date<br/>dim_service_item<br/>dim_agent<br/>dim_group<br/>dim_requester"]:::dim
            INT["int_conversation_event<br/>規則：持球（誰在等誰）"]:::fact
            FACT["fact_ticket<br/>fact_conversation<br/>fact_agent_touch"]:::fact
            DQ["dq_unmatched<br/>對不到清冊的值"]:::warn
        end
        subgraph L3["L3 輸出｜30_mart.sql"]
            CREDIT["mart_agent_credit<br/>規則：點數怎麼分"]:::mart
            LONG["mart_opbrain_fresh_long<br/>人效長表"]:::mart
        end
        META["mart_param<br/>build_log<br/>build_info"]:::meta
    end

    subgraph USE["使用者"]
        OPB["OpBrain 人效總攬<br/>Email / Fresh total"]:::user
        OPA["OP 分析<br/>處理時長、持球"]:::user
        OWNER["清冊維護者"]:::user
    end

    FS -->|fresh_export.py| RAW
    XLSX -->|手動複製| REFX
    RAW -->|ATTACH 唯讀| STG
    REFX -->|openpyxl| REF
    REF --> STG
    STG --> DIM
    STG --> INT
    DIM --> FACT
    INT --> FACT
    DIM --> DQ
    FACT --> CREDIT
    CREDIT --> LONG
    DIM --> LONG
    LONG --> OPB
    FACT --> OPA
    DQ --> OWNER

    classDef src  fill:#FAFAFA,stroke:#9E9E9E,color:#212121,stroke-dasharray:4 3
    classDef raw  fill:#ECEFF1,stroke:#546E7A,color:#263238
    classDef stg  fill:#E0F2F1,stroke:#00897B,color:#004D40
    classDef dim  fill:#EDE7F6,stroke:#5E35B1,color:#311B92
    classDef fact fill:#E3F2FD,stroke:#1565C0,color:#0D47A1,stroke-width:2px
    classDef mart fill:#EFEBE9,stroke:#6D4C41,color:#3E2723,stroke-width:2px
    classDef meta fill:#F5F5F5,stroke:#BDBDBD,color:#616161,stroke-dasharray:2 2
    classDef user fill:#FFFFFF,stroke:#424242,color:#212121
    classDef warn fill:#FFF8E1,stroke:#F9A825,color:#5D4037

    style SRC  fill:#FFFFFF,stroke:#BDBDBD,color:#424242
    style L0   fill:#F7F9FA,stroke:#546E7A,color:#263238
    style MART fill:#FFFFFF,stroke:#424242,color:#212121
    style L1   fill:#F3FAF9,stroke:#00897B,color:#004D40
    style L2   fill:#F8F6FC,stroke:#5E35B1,color:#311B92
    style L3   fill:#F9F6F5,stroke:#6D4C41,color:#3E2723
    style USE  fill:#FFFFFF,stroke:#BDBDBD,color:#424242
```

| 層 | SQL 檔 | 放什麼規則 |
|---|---|---|
| L0 原始 | — | 不修改；清冊只複製需要的欄位（不含 Email） |
| L1 標準化 | `10_stg.sql` | 不放業務規則，只改名、合併、正規化 join key |
| L2 核心 | `20_core_dims.sql`、`21_core_facts.sql` | 穩定的規則：持球、計分口徑（v1／v2）、姓名對照 |
| L3 輸出 | `30_mart.sql` | 會變的政策：點數分配（`credit_rule`） |

---

## 2. Table Roles (表的分工：維度 vs 事實)

先看這張再看第 3 節的欄位細節。**紫色的維度表圍在外圈**，提供「從哪個角度看」；**藍色的事實表在中間**，存「發生了什麼、多少」；**棕色的輸出表**由事實表彙總而來。
箭頭上的文字是 join 用的 key。

```mermaid
flowchart LR
    DD["dim_date<br/>日期、週、月、是否週末"]:::dim
    DS["dim_service_item<br/>服務項目、負責團隊<br/>人效第二層、點數"]:::dim
    DG["dim_group<br/>處理群組、監管、DP/WD/AC"]:::dim
    DR["dim_requester<br/>申請人、區域、市場"]:::dim
    DA["dim_agent<br/>經手人、Office、Team、Role"]:::dim

    FT["fact_ticket<br/>一張工單一列<br/>時長拆解、點數"]:::fact
    FC["fact_conversation<br/>一則對話一列<br/>這段在等誰、等多久"]:::fact
    FA["fact_agent_touch<br/>工單 × 經手人<br/>回覆幾則、是否結案"]:::fact

    MC["mart_agent_credit<br/>工單 × 經手人<br/>分到幾點"]:::mart
    ML["mart_opbrain_fresh_long<br/>日 × 項目 × 人<br/>人效長表"]:::mart

    DS -->|service_item_key| FT
    DG -->|group_key| FT
    DR -->|requester_key| FT
    DD -->|created / resolved_date_key| FT
    FT -->|ticket_id| FC
    FT -->|ticket_id| FA
    DA -->|agent_key| FC
    DA -->|agent_key| FA
    FA -->|套用 credit_rule| MC
    FT -->|工單點數| MC
    MC -->|彙總| ML

    classDef dim  fill:#EDE7F6,stroke:#5E35B1,color:#311B92
    classDef fact fill:#E3F2FD,stroke:#1565C0,color:#0D47A1,stroke-width:2px
    classDef mart fill:#EFEBE9,stroke:#6D4C41,color:#3E2723,stroke-width:2px
```

讀法（由上往下鑽）：

```text
人效第二層 Sub category（dim_service_item）
  └─ 服務項目 item
       └─ 工單 fact_ticket：時長拆成 OP 持球 / 等申請方
            └─ 對話 fact_conversation：每段等待算給誰
                 └─ 經手人 fact_agent_touch -> 分到的點數 mart_agent_credit（Office、Team 來自 dim_agent）
```

---

## 3. Star Schema Detail (星狀關聯：欄位細節)

和第 2 節是同一組表，這裡列出主要欄位與 key（ER 圖在多數預覽器不支援上色，顏色請對照第 2 節）。
`*_key = -1` 代表值本身是空的（未填服務項目、無人經手等）；對不到清冊的值會保留成維度列，不會丟掉工單。

```mermaid
erDiagram
    dim_date {
        int date_key PK
        text date
        text year_month
        text week_start
        int is_weekend
    }
    dim_service_item {
        int service_item_key PK
        text item_key "正規化自然鍵"
        text item
        int in_roster
        text owner_team "負責團隊"
        text sub_category_for_op "人效第二層"
        real points_for_op
        text handling_type "OP 自辦 / OP 轉派 / 其他部門"
        text workspace
        text fresh_mapping
    }
    dim_agent {
        int agent_key PK
        text author_key "正規化自然鍵"
        text author_name
        int is_system
        int in_roster
        text match_method
        text roster_team
        text roster_office
        text roster_role
    }
    dim_group {
        int group_key PK
        text group_name
        text regulator
        text function "DP / WD / AC / Other"
    }
    dim_requester {
        int requester_key PK
        text requester
        text department_name
        text region
        text market
        text team
    }
    fact_ticket {
        int ticket_id PK
        int created_date_key FK
        int resolved_date_key FK
        int service_item_key FK
        int group_key FK
        int requester_key FK
        real total_hours
        real op_hold_hours
        real requester_wait_hours
        real first_response_hours
        int n_agents
        text scoring_version "v1 / v2"
        real points_for_op
    }
    fact_conversation {
        int ticket_id PK, FK
        int conversation_id PK
        int turn_seq
        int created_date_key FK
        int agent_key FK
        text event_type "agent_public / requester_reply / note / system"
        text waiting_on "OP / REQUESTER"
        real wait_before_hours
    }
    fact_agent_touch {
        int ticket_id PK, FK
        int agent_key PK, FK
        int n_public_replies
        int n_notes
        int is_first_responder
        int is_last_public_replier
    }
    mart_agent_credit {
        int ticket_id PK, FK
        int agent_key PK, FK
        text credit_rule
        real credit_share
        real credited_points
        int credit_date_key FK
    }
    mart_opbrain_fresh_long {
        text work_date
        text layer2
        text item
        text agent_name
        int tickets_touched
        real credited_points
    }

    dim_date          ||--o{ fact_ticket        : "created / resolved"
    dim_service_item  ||--o{ fact_ticket        : "service_item_key"
    dim_group         ||--o{ fact_ticket        : "group_key"
    dim_requester     ||--o{ fact_ticket        : "requester_key"
    fact_ticket       ||--o{ fact_conversation  : "ticket_id"
    dim_agent         ||--o{ fact_conversation  : "agent_key"
    dim_date          ||--o{ fact_conversation  : "created_date_key"
    fact_ticket       ||--o{ fact_agent_touch   : "ticket_id"
    dim_agent         ||--o{ fact_agent_touch   : "agent_key"
    fact_agent_touch  ||--|| mart_agent_credit  : "套用 credit_rule"
    fact_ticket       ||--|{ mart_agent_credit  : "ticket points"
    dim_date          ||--o{ mart_agent_credit  : "credit_date_key"
    mart_agent_credit }o--|| mart_opbrain_fresh_long : "日 x 項目 x 人 彙總"
```

---

## 4. Build Sequence (一次建置的步驟)

步驟編號與 `mart_build.py` 的 `# Step N` 一致。**藍色**是一般步驟，**綠色**是成功結束，**紅色**是失敗（mart.db 維持上一版），**黃色**是建置仍成功但需要補資料。

```mermaid
flowchart TD
    S1["Step 1 檢查參數<br/>credit_rule、scoring_v2_from"]:::step --> S2
    S2["Step 2 讀清冊快照<br/>工作表、欄位、Item 重複"]:::step --> S3
    S3{"Step 3 找到 db/*.db？"}:::step
    S3 -- 沒有 --> END0(["INFO 結束，exit 0<br/>沒有工作不算錯誤"]):::meta
    S3 -- 有 --> S4["Step 4 建立 mart.db.building<br/>ATTACH 唯讀、寫入 ref_*、mart_param"]:::step
    S4 --> S5["Step 5 執行 SQL<br/>10_stg → 20_core_dims → 21_core_facts → 30_mart"]:::step
    S5 --> S6{"Step 6 驗收<br/>工單數、對話數、時長、點數加總"}:::step
    S6 -- 未通過 --> FAIL(["ERROR：刪除 .building<br/>mart.db 維持上一版"]):::err
    S6 -- 通過 --> S7["Step 7 原子取代<br/>mart.db.building → mart.db"]:::step
    S7 --> OK(["完成，結果寫入 build_info"]):::ok
    S6 -.->|"對照率 &lt; 95%"| WARN["WARNING：請看 dq_unmatched<br/>建置仍完成"]:::warn

    classDef step fill:#E3F2FD,stroke:#1565C0,color:#0D47A1
    classDef meta fill:#F5F5F5,stroke:#BDBDBD,color:#616161,stroke-dasharray:2 2
    classDef warn fill:#FFF8E1,stroke:#F9A825,color:#5D4037
    classDef ok   fill:#E8F5E9,stroke:#2E7D32,color:#1B5E20,stroke-width:2px
    classDef err  fill:#FFEBEE,stroke:#C62828,color:#B71C1C,stroke-width:2px
```

---

## 5. Packaging (打包帶走)

| 要做什麼 | 帶走哪些 | 說明 |
|---|---|---|
| **只用資料**（查詢、分析、接 BI） | `mart/mart.db` 一個檔案 | 已完全實體化：沒有 view、trigger，也不依賴原始 .db；所有層、清冊快照、參數和驗收結果都在裡面。用任何 SQLite 工具都能直接開 |
| **在別台機器重建** | 程式（`mart_build.py`、`sql/mart/`、`fresh_export.py`、`freshservice_api.py`、`db_query.py`、`ticket_db.py`、`logger.py`、`config.py`、`requirements.txt`）＋ `db/` ＋ `ref/` | `mart_build.py` 會 import 這些既有模組 |

> [!CAUTION]
> mart.db 含個資（客服人名、申請人、工單主旨、清冊的 Roster 資訊），請以內部資料的等級傳遞，不要放進 repo 或作品集。
> `config.py` 目前寫死了 secret，打包程式前請先輪替並改成從環境變數讀取。
