# vba-mcp-server

[English](#english) | [日本語](#japanese)

---

<a name="english"></a>

## Overview

An MCP (Model Context Protocol) server that enables AI coding assistants like Claude Desktop and Cursor to read and write VBA code in Excel and Access files.

**Transform your VBA development experience** - No more copy-pasting code between your IDE and Office applications. Edit VBA code directly from your AI-powered development environment.

## Features

- 📖 **Read VBA modules** - List and read code from any VBA module
- ✏️ **Write VBA modules** - Update or create VBA code programmatically
- 📦 **Export/Import** - Export modules to files and import them back
- 🔒 **Auto-backup** - Automatic backup before any write operation
- 🔍 **Procedure-level access** - Read and write individual procedures

### Supported Module Types

| Type | Read | Write | Notes |
|------|------|-------|-------|
| Standard Module (.bas) | ✅ | ✅ | Full support |
| Class Module (.cls) | ✅ | ✅ | Full support |
| UserForm (.frm) | ✅ | ✅ | Code only, not design |
| Document Module | ✅ | ✅ | ThisWorkbook, Sheet modules |
| Access Form/Report | ✅ | ✅ | Code-behind only |

## Quick Start

### Prerequisites

1. Windows 10/11
2. Microsoft Office 2016 or later (including Microsoft 365)
3. Enable "Trust access to the VBA project object model" in Office settings

### Installation

#### Option 1: Using Installer (Recommended)

1. Download `VbaMcpServer.msi` from [Releases](../../releases) page
2. Run the installer and follow the wizard
3. Launch "VBA MCP Server Manager" from Start Menu

#### Option 2: Build from Source

```bash
git clone https://github.com/YOUR_USERNAME/vba-mcp-server.git
cd vba-mcp-server

# Build all projects (outputs to unified bin/Release/ directory)
dotnet build -c Release

# Or build self-contained single executables
dotnet publish src/VbaMcpServer -c Release -r win-x64 --self-contained /p:PublishSingleFile=true
dotnet publish src/VbaMcpServer.GUI -c Release -r win-x64 --self-contained /p:PublishSingleFile=true
```

**Output locations:**
- Normal build: `bin/Release/` (all executables in one directory)
- Publish build: `src/{ProjectName}/bin/Release/win-x64/publish/`

### Configuration

#### Using GUI Manager

1. Launch "VBA MCP Server Manager" from Start Menu
2. The GUI automatically detects VbaMcpServer.exe using:
   - Registry entry (set by installer)
   - Same directory location
   - Development build detection
3. (Optional) Override path in `appsettings.json` if needed
4. Click "Start" to run the MCP server
5. Monitor logs in real-time

For detailed configuration options, see [docs/CONFIGURATION.md](docs/CONFIGURATION.md).

#### Manual Configuration (CLI)

Add to your Claude Desktop config (`%APPDATA%\Claude\claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "vba": {
      "command": "C:\\Program Files\\VBA MCP Server\\VbaMcpServer.exe"
    }
  }
}
```

Or if you built from source:

```json
{
  "mcpServers": {
    "vba": {
      "command": "C:\\path\\to\\VbaMcpServer.exe"
    }
  }
}
```

## Usage Examples

Once configured, you can ask Claude to:

- "List all VBA modules in C:\Projects\MyWorkbook.xlsm"
- "Show me the code in Module1"
- "Add error handling to the SaveData procedure"
- "Create a new class module called DataProcessor"
- "Refactor this code to use early binding"

## Office Security Settings

⚠️ **Required Setting**: You must enable VBA project access in Office:

1. Open Excel or Access
2. Go to **File** → **Options** → **Trust Center**
3. Click **Trust Center Settings**
4. Select **Macro Settings**
5. Check ✅ **Trust access to the VBA project object model**

See [docs/SECURITY.md](docs/SECURITY.md) for detailed instructions.

## Available Tools

| Tool | Description |
|------|-------------|
| `list_open_files` | List currently open Office files |
| `list_modules` | List all VBA modules in a file |
| `read_module` | Read entire module code |
| `write_module` | Write/replace module code |
| `create_module` | Create a new module |
| `delete_module` | Delete a module |
| `list_procedures` | List procedures in a module |
| `read_procedure` | Read a specific procedure |
| `export_module` | Export module to file |
| `import_module` | Import module from file |

## Building from Source

### Requirements

- .NET 8 SDK or later
- Visual Studio 2022 or VS Code with C# extension

### Build

```bash
cd src/VbaMcpServer
dotnet build
```

### Test

```bash
cd tests/VbaMcpServer.Tests
dotnet test
```

### Publish

```bash
dotnet publish -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true
```

## Contributing

Contributions are welcome! Please read [CONTRIBUTING.md](CONTRIBUTING.md) before submitting PRs.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

<a name="japanese"></a>

## 概要

Excel や Access の VBA コードを、Claude Desktop や Cursor などの AI コーディング環境から直接読み書きできるようにする MCP（Model Context Protocol）サーバーです。

**VBA 開発体験を一新** - Office アプリケーションと IDE 間でのコードのコピー＆ペーストが不要に。AI 搭載の開発環境から直接 VBA コードを編集できます。

## 機能

- 📖 **VBA モジュールの読み取り** - すべての VBA モジュールの一覧表示とコード取得
- ✏️ **VBA モジュールの書き込み** - プログラムからの VBA コード更新・作成
- 📦 **エクスポート/インポート** - モジュールのファイル出力と読み込み
- 🔒 **自動バックアップ** - 書き込み操作前の自動バックアップ
- 🔍 **プロシージャ単位のアクセス** - 個別のプロシージャの読み書き

### 対応モジュールタイプ

| タイプ | 読み取り | 書き込み | 備考 |
|--------|---------|---------|------|
| 標準モジュール (.bas) | ✅ | ✅ | 完全対応 |
| クラスモジュール (.cls) | ✅ | ✅ | 完全対応 |
| ユーザーフォーム (.frm) | ✅ | ✅ | コードのみ、デザインは不可 |
| ドキュメントモジュール | ✅ | ✅ | ThisWorkbook、Sheet モジュール |
| Access フォーム/レポート | ✅ | ✅ | コードビハインドのみ |

## クイックスタート

### 前提条件

1. Windows 10/11
2. Microsoft Office 2016 以降（Microsoft 365 含む）
3. Office の設定で「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を有効化

### インストール

#### 方法1: インストーラを使用（推奨）

1. [Releases](../../releases) ページから `VbaMcpServer.msi` をダウンロード
2. インストーラを実行してウィザードに従う
3. スタートメニューから「VBA MCP Server Manager」を起動

#### 方法2: ソースからビルド

```bash
git clone https://github.com/YOUR_USERNAME/vba-mcp-server.git
cd vba-mcp-server

# 全プロジェクトをビルド（統一された bin/Release/ ディレクトリに出力）
dotnet build -c Release

# または、自己完結型の単一実行ファイルをビルド
dotnet publish src/VbaMcpServer -c Release -r win-x64 --self-contained /p:PublishSingleFile=true
dotnet publish src/VbaMcpServer.GUI -c Release -r win-x64 --self-contained /p:PublishSingleFile=true
```

**出力先:**
- 通常ビルド: `bin/Release/` (すべての実行ファイルが同じディレクトリ)
- Publishビルド: `src/{ProjectName}/bin/Release/win-x64/publish/`

### 設定

#### GUI マネージャーを使用

1. スタートメニューから「VBA MCP Server Manager」を起動
2. GUI は以下の方法で VbaMcpServer.exe を自動検出します:
   - レジストリエントリ（インストーラーで設定）
   - 同じディレクトリ内の検索
   - 開発ビルドの自動検出
3. （オプション）必要に応じて `appsettings.json` でパスを上書き可能
4. 「Start」ボタンをクリックして MCP サーバーを実行
5. リアルタイムでログを監視

詳細な設定オプションは [docs/CONFIGURATION.md](docs/CONFIGURATION.md) を参照してください。

#### 手動設定（CLI）

Claude Desktop の設定ファイル（`%APPDATA%\Claude\claude_desktop_config.json`）に追加：

```json
{
  "mcpServers": {
    "vba": {
      "command": "C:\\Program Files\\VBA MCP Server\\VbaMcpServer.exe"
    }
  }
}
```

またはソースからビルドした場合：

```json
{
  "mcpServers": {
    "vba": {
      "command": "C:\\path\\to\\VbaMcpServer.exe"
    }
  }
}
```

## 使用例

設定後、Claude に以下のように依頼できます：

- 「C:\Projects\MyWorkbook.xlsm の VBA モジュール一覧を表示して」
- 「Module1 のコードを見せて」
- 「SaveData プロシージャにエラーハンドリングを追加して」
- 「DataProcessor という新しいクラスモジュールを作成して」
- 「このコードを事前バインディングを使うようにリファクタリングして」

## Office セキュリティ設定

⚠️ **必須設定**: Office で VBA プロジェクトへのアクセスを有効にする必要があります：

1. Excel または Access を開く
2. **ファイル** → **オプション** → **トラストセンター** を選択
3. **トラストセンターの設定** をクリック
4. **マクロの設定** を選択
5. ✅ **VBA プロジェクト オブジェクト モデルへのアクセスを信頼する** にチェック

詳細は [docs/SECURITY.md](docs/SECURITY.md) を参照してください。

## ライセンス

MIT License - 詳細は [LICENSE](LICENSE) ファイルを参照してください。
