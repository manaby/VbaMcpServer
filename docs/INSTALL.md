# Installation Guide / インストールガイド

[English](#english) | [日本語](#japanese)

---

<a name="english"></a>

## System Requirements

- Windows 10 or Windows 11
- Microsoft Office 2016 or later (including Microsoft 365)
- .NET 8 Runtime (included in self-contained builds)

## Installation Options

### Option 1: Download Pre-built Binary (Recommended)

1. Go to [Releases](../../releases)
2. Download the latest `VbaMcpServer.exe`
3. Place it in a permanent location (e.g., `C:\Program Files\VbaMcpServer\`)
4. Configure your MCP client (see below)

### Option 2: Build from Source

#### Prerequisites

- .NET 8 SDK or later
- Visual Studio 2022 or Visual Studio Code with C# extension
- Microsoft Office installed (for COM references)

#### Steps

```bash
# Clone the repository
git clone https://github.com/YOUR_USERNAME/vba-mcp-server.git
cd vba-mcp-server

# Build
cd src/VbaMcpServer
dotnet build

# Or publish as single executable
dotnet publish -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true
```

The executable will be in `bin/Release/net8.0-windows/win-x64/publish/`.

## Configuration

### Claude Desktop

Edit `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "vba": {
      "command": "C:\\Program Files\\VbaMcpServer\\VbaMcpServer.exe"
    }
  }
}
```

### Cursor

Add to your Cursor MCP settings:

```json
{
  "mcpServers": {
    "vba": {
      "command": "C:\\Program Files\\VbaMcpServer\\VbaMcpServer.exe"
    }
  }
}
```

### VS Code with Continue Extension

Add to your Continue configuration:

```json
{
  "models": [...],
  "mcpServers": [
    {
      "name": "vba",
      "command": "C:\\Program Files\\VbaMcpServer\\VbaMcpServer.exe"
    }
  ]
}
```

## Post-Installation

### Enable VBA Project Access

**This step is required!**

See [SECURITY.md](SECURITY.md) for detailed instructions on enabling "Trust access to the VBA project object model" in Office.

### Verify Installation

1. Open an Excel file with macros (.xlsm)
2. In your MCP client, ask: "List open Excel files"
3. You should see your workbook listed

## Troubleshooting

### "Excel is not running" error

Make sure Excel is open with a workbook before attempting to use the VBA tools.

### "Workbook not found" error

The workbook must be open in Excel. Check:
- The file path is correct
- The file is actually open in Excel
- Use the full path (e.g., `C:\Projects\MyWorkbook.xlsm`)

### "VBA project access is not trusted" error

Enable VBA project access in Office Trust Center. See [SECURITY.md](SECURITY.md).

### MCP server not connecting

1. Check that the executable path in your config is correct
2. Ensure the path uses double backslashes (`\\`) in JSON
3. Restart your MCP client after changing configuration

---

<a name="japanese"></a>

## システム要件

- Windows 10 または Windows 11
- Microsoft Office 2016 以降（Microsoft 365 含む）
- .NET 8 ランタイム（self-contained ビルドには同梱）

## インストール方法

### 方法 1: ビルド済みバイナリをダウンロード（推奨）

1. [Releases](../../releases) にアクセス
2. 最新の `VbaMcpServer.exe` をダウンロード
3. 固定の場所に配置（例: `C:\Program Files\VbaMcpServer\`）
4. MCP クライアントを設定（下記参照）

### 方法 2: ソースからビルド

#### 前提条件

- .NET 8 SDK 以降
- Visual Studio 2022 または Visual Studio Code（C# 拡張機能付き）
- Microsoft Office がインストールされていること（COM 参照用）

#### 手順

```bash
# リポジトリをクローン
git clone https://github.com/YOUR_USERNAME/vba-mcp-server.git
cd vba-mcp-server

# ビルド
cd src/VbaMcpServer
dotnet build

# または単一実行ファイルとして発行
dotnet publish -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true
```

実行ファイルは `bin/Release/net8.0-windows/win-x64/publish/` に出力されます。

## 設定

### Claude Desktop

`%APPDATA%\Claude\claude_desktop_config.json` を編集：

```json
{
  "mcpServers": {
    "vba": {
      "command": "C:\\Program Files\\VbaMcpServer\\VbaMcpServer.exe"
    }
  }
}
```

### Cursor

Cursor の MCP 設定に追加：

```json
{
  "mcpServers": {
    "vba": {
      "command": "C:\\Program Files\\VbaMcpServer\\VbaMcpServer.exe"
    }
  }
}
```

## インストール後の設定

### VBA プロジェクトアクセスを有効にする

**この手順は必須です！**

Office で「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を有効にする詳細な手順は [SECURITY.md](SECURITY.md) を参照してください。

### インストールの確認

1. マクロを含む Excel ファイル（.xlsm）を開く
2. MCP クライアントで「開いている Excel ファイルを一覧表示して」と尋ねる
3. ワークブックが一覧に表示されれば成功

## トラブルシューティング

### 「Excel が実行されていません」エラー

VBA ツールを使用する前に、Excel でワークブックを開いていることを確認してください。

### 「ワークブックが見つかりません」エラー

ワークブックが Excel で開いている必要があります。以下を確認：
- ファイルパスが正しいこと
- ファイルが実際に Excel で開いていること
- フルパスを使用すること（例: `C:\Projects\MyWorkbook.xlsm`）

### 「VBA プロジェクトへのアクセスが信頼されていません」エラー

Office トラストセンターで VBA プロジェクトアクセスを有効にしてください。[SECURITY.md](SECURITY.md) を参照。

### MCP サーバーに接続できない

1. 設定ファイル内の実行ファイルパスが正しいことを確認
2. JSON 内でバックスラッシュがダブル（`\\`）になっていることを確認
3. 設定変更後に MCP クライアントを再起動
