# Outlook Web MCP Server（非公式）

Outlook Web が内部で利用している非公開 API を使って、メールと予定表を MCP クライアントから扱うためのローカル MCP サーバーです。Microsoft Graph、Outlook デスクトップ、COM、ブラウザー Cookie は使いません。デスクトップ版の COM 連携が使えない環境で、その代替として動作します。

> [!WARNING]
> 接続先の `https://outlook.cloud.microsoft/owa/service.svc` は、Microsoft の公開・サポート対象 API ではありません。Outlook Web 側の変更や組織ポリシーにより、予告なく動かなくなる可能性があります。本番の基幹連携には Microsoft Graph を推奨します。

## 特徴

- COM 版（outlook-com-mcp-server）とツール名・主要な引数・戻り値が互換
- メールの検索・本文取得、フォルダー一覧、予定の取得
- 返信／全員返信の下書き保存と、既読・未読の変更
- 認証状態の確認ツール `outlook_auth_status`

送信、削除、移動、アーカイブのツールは意図的に実装していません。

| ツール | 動作 | メールボックス変更 |
|---|---|---|
| `search_emails` | 件名・送信者を検索 | なし |
| `list_mail_folders` | フォルダー一覧 | なし |
| `get_email` | 1通の本文を取得 | なし |
| `set_email_read_state` | 既読・未読を変更 | あり |
| `list_calendar_events` | 予定を期間指定で取得 | なし |
| `create_reply_draft` | 返信／全員返信の下書きを保存 | 下書き追加のみ |
| `outlook_auth_status` | 認証状態を確認 | なし |

COM 版との違いとして、`storeId` は COM の Store Entry ID ではなく、このサーバー専用の値 `outlook-web` です。検索結果から返された値をそのまま使ってください。また、COM 版で取得した `emailId`、`folderId`、`storeId` は再利用できません。切り替え時は、検索とフォルダー一覧から ID を取り直してください。

## 動作要件

- Python 3.11 以上
- [uv](https://docs.astral.sh/uv/)

## セットアップ

```powershell
uv sync --dev
uv run outlook-web-mcp login
uv run outlook-web-mcp probe
```

初回ログインでは Microsoft のデバイスコード画面が開きます。アクセストークンはメモリ内だけで扱い、更新トークンは OS の資格情報保管庫に保存します。標準出力や MCP 応答にトークンを出すことはありません。

サインインを別工程に分けたい場合は、発行と完了を分離できます。

```powershell
uv run outlook-web-mcp login-start
# 表示されたコードでサインイン後
uv run outlook-web-mcp login-complete
```

## 使い方

stdio サーバーの起動コマンドは次です。

```powershell
uv run outlook-web-mcp-server
```

MCP クライアントへの登録例（パスはこのリポジトリの絶対パスへ変更してください）:

```json
{
  "mcpServers": {
    "outlook-web": {
      "command": "uv",
      "args": [
        "--directory",
        "C:\\path\\to\\outlook-web-mcp-server",
        "run",
        "outlook-web-mcp-server"
      ]
    }
  }
}
```

COM 版から切り替える手順:

1. `uv run outlook-web-mcp login` で、このサーバー専用の認証を1回行う。
2. MCP クライアントの `outlook-com` 起動設定を、上記の `outlook-web` 設定へ置き換える。
3. MCP クライアントを再起動し、`outlook_auth_status` と `search_emails` で確認する。

### CLI

認証や疎通確認のためのサブコマンドを用意しています。

```text
outlook-web-mcp login          # 対話ログイン
outlook-web-mcp login-start    # デバイスコードを発行
outlook-web-mcp login-complete # 発行済みコードのログインを完了
outlook-web-mcp status         # 秘密を含まない認証状態
outlook-web-mcp probe          # メール内容を取得しない Inbox 疎通確認
outlook-web-mcp logout         # 保存済み資格情報を削除
outlook-web-mcp serve          # stdio MCP を起動
```

## 設定

必要な場合だけ、環境変数で既定値を上書きします。

| 変数 | 既定値 |
|---|---|
| `OUTLOOK_WEB_CREDENTIAL_PROFILE` | `default` |
| `OUTLOOK_WEB_TIME_ZONE_ID` | `Tokyo Standard Time` |
| `OUTLOOK_WEB_REQUEST_TIMEOUT_SECONDS` | `30` |
| `OUTLOOK_WEB_MAX_RESPONSE_BYTES` | `10000000` |
| `OUTLOOK_WEB_MAX_SCANNED_ITEMS` | `5000` |
| `OUTLOOK_WEB_MAX_SCANNED_FOLDERS` | `100` |
| `OUTLOOK_WEB_LOG_LEVEL` | `INFO` |

`OUTLOOK_WEB_CLIENT_ID` も変更できますが、既定値は Outlook Web が公開ログイン画面で使うファーストパーティーのクライアント ID です。組織の条件付きアクセスや同意ポリシーによっては認証できない場合があります。

## セキュリティと制約

- 送信、削除、移動、アーカイブはできません。書き込みは下書きの追加と既読・未読の変更だけです。
- アクセストークンはメモリ内のみ、更新トークンは OS の資格情報保管庫に保存し、ログや応答に出しません。
- 接続先は非公開 API のため、Outlook Web 側の変更で予告なく動かなくなる可能性があります（冒頭の警告を参照）。

## 開発

```powershell
uv run pytest
uv run ruff check .
uv run mypy
uv build
```

内部 API の調査結果と設計上の境界は [docs/internal-api.md](docs/internal-api.md) にあります。

## ライセンス

MIT License の下で公開しています。詳細は [LICENSE](LICENSE) を参照してください。

Copyright (c) 2026 ma-nakaya
