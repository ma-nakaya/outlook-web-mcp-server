# Outlook Web MCP server (unofficial)

ローカルの Outlook COM を、Outlook Web が内部利用している非公開 API に置き換える
MCP サーバーです。Microsoft Graph、Outlook デスクトップ、COM、ブラウザー Cookie は
使いません。

> [!WARNING]
> `https://outlook.cloud.microsoft/owa/service.svc` は Microsoft の公開・サポート対象 API
> ではありません。Outlook Web 側の変更や組織ポリシーで、予告なく動かなくなる可能性が
> あります。本番の基幹連携には Microsoft Graph を推奨します。

## COM 版との互換性

COM 版のツール名と主要な引数・戻り値を維持しています。

| ツール | 動作 | メールボックス変更 |
|---|---|---|
| `search_emails` | 件名・送信者を検索 | なし |
| `list_mail_folders` | フォルダー一覧 | なし |
| `get_email` | 1通の本文を取得 | なし |
| `set_email_read_state` | 既読・未読を変更 | あり |
| `list_calendar_events` | 予定を期間指定で取得 | なし |
| `create_reply_draft` | 返信／全員返信の下書きを保存 | 下書き追加のみ |

認証状態確認用に `outlook_auth_status` も追加しています。送信、削除、移動、アーカイブの
ツールは意図的に実装していません。`storeId` は COM の Store Entry ID ではなく、
このサーバー専用の `outlook-web` です。検索結果から返された値をそのまま使ってください。

## セットアップ

Python 3.11 以上と [uv](https://docs.astral.sh/uv/) が必要です。

```powershell
uv sync --dev
uv run outlook-web-mcp login
uv run outlook-web-mcp probe
```

初回ログインでは Microsoft のデバイスコード画面が開きます。アクセストークンはメモリ内
だけで扱い、更新トークンは OS の資格情報保管庫に保存します。標準出力や MCP 応答に
トークンを出しません。

サインインを別工程に分けたい場合:

```powershell
uv run outlook-web-mcp login-start
# 表示されたコードでサインイン後
uv run outlook-web-mcp login-complete
```

## MCP への登録

stdio サーバーの起動コマンドは次です。

```powershell
uv run outlook-web-mcp-server
```

クライアント設定例（パスはこのリポジトリの絶対パスへ変更）:

```json
{
  "mcpServers": {
    "outlook-web": {
      "command": "uv",
      "args": [
        "--directory",
        "C:\\path\\to\\outlook-web-teams-un-oficial-api",
        "run",
        "outlook-web-mcp-server"
      ]
    }
  }
}
```

### COM 版から切り替える

1. `uv run outlook-web-mcp login` でこのサーバー専用の認証を1回行う。
2. MCP クライアントの `outlook-com` 起動設定を、上記 `outlook-web` の設定へ置き換える。
3. MCP クライアントを再起動し、`outlook_auth_status` と `search_emails` で確認する。

ツール名と引数名は維持していますが、COM 版で以前取得した `emailId`、`folderId`、
`storeId` は再利用できません。Web 版の検索・フォルダー一覧から ID を取り直してください。

## CLI

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

必要な場合だけ環境変数で上書きします。

| 変数 | 既定値 |
|---|---|
| `OUTLOOK_WEB_CREDENTIAL_PROFILE` | `default` |
| `OUTLOOK_WEB_TIME_ZONE_ID` | `Tokyo Standard Time` |
| `OUTLOOK_WEB_REQUEST_TIMEOUT_SECONDS` | `30` |
| `OUTLOOK_WEB_MAX_RESPONSE_BYTES` | `10000000` |
| `OUTLOOK_WEB_MAX_SCANNED_ITEMS` | `5000` |
| `OUTLOOK_WEB_MAX_SCANNED_FOLDERS` | `100` |
| `OUTLOOK_WEB_LOG_LEVEL` | `INFO` |

`OUTLOOK_WEB_CLIENT_ID` も変更できますが、既定値は Outlook Web が公開ログイン画面で使う
ファーストパーティー クライアント ID です。組織の条件付きアクセスや同意ポリシーによって
認証できない場合があります。

## 開発時の確認

```powershell
uv run pytest
uv run ruff check .
uv run mypy
uv build
```

内部 API の調査結果と設計上の境界は [docs/internal-api.md](docs/internal-api.md) にあります。
