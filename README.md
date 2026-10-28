# タスク娘 (Tasuku-Musume) - AI駆動型ゲーミフィケーションタスク管理Bot

Discord上でタスクの進捗を報告することで、AIキャラクターと対話しながらモチベーションを維持できるゲーミフィケーションタスク管理Botです。ローコードツールと生成AIを組み合わせ、完全サーバーレスで構築しています。

## システムアーキテクチャ

```mermaid
flowchart LR
    User([ユーザー]) -- タスク報告 --> Discord[Discord Bot]
    Discord -- Webhook --> Make[Make]

    Make -- 情報取得/更新 --> GAS[Google Apps Script]
    GAS <-- データ読み書き --> Sheets[(Google Sheets\n好感度・設定)]

    Make -- プロンプト送信 --> Groq[Groq API]
    Groq -- セリフ生成 --> Make

    Make -- メッセージ送信 --> Discord
    Discord -- リプライ --> User
