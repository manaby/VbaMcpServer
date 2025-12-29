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
2. **Click "Browse" button to select your target Excel/Access file**
3. Click "Start" to run the MCP server
4. Monitor logs in real-time

**Notes:**
- The GUI automatically detects VbaMcpServer.exe using registry entry (set by installer) or same directory location
- You can override the server path in `appsettings.json` if needed
- The selected target file will be automatically opened when the server starts

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

Or for Claude Code (CLI tool):

**Windows:**
```json
{
  "mcpServers": {
    "vba": {
      "command": "C:\\Program Files\\VBA MCP Server\\VbaMcpServer.exe"
    }
  }
}
```

**macOS/Linux:**
```json
{
  "mcpServers": {
    "vba": {
      "command": "/path/to/VbaMcpServer.exe"
    }
  }
}
```

Configuration file location:
- Windows: `%USERPROFILE%\.claude\settings.json`
- macOS/Linux: `~/.claude/settings.json`

## ⚠️ Important: Backup and Version Control

**This tool does NOT provide automatic backup functionality.** VBA code changes are irreversible operations. You are responsible for protecting your work:

### Recommended Practices

1. **Use Git for VBA Code**: Manage your VBA code with Git or other version control systems
2. **Backup Files Before Editing**: Always create a copy of your Excel/Access file before making code changes
3. **Use Office AutoSave**: If using OneDrive/SharePoint, leverage the automatic version history feature

**VBA code modifications are permanent and cannot be undone by this tool. Always backup your files before making changes.**

## Usage Examples

Once configured, you can ask Claude to:

### Excel Examples

- "List all VBA modules in C:\Projects\MyWorkbook.xlsm"
- "Show me the code in Module1"
- "Add error handling to the SaveData procedure"
- "Create a new class module called DataProcessor"
- "Refactor this code to use early binding"

### Access Examples

- "List all VBA modules in C:\Projects\MyDatabase.accdb"
- "Show me the code in the Form_MainForm module"
- "Add error handling to the btnSave_Click procedure in Form_MainForm"
- "Create a new class module called DatabaseConnection in the Access database"

### Access Data Examples

- "List all tables in C:\Projects\MyDatabase.accdb"
- "Show me the structure of the Customers table"
- "Get the first 50 records from the Orders table where OrderDate > #2024-01-01#"
- "List all queries in the database"
- "Show me the SQL for the qryMonthlyReport query"
- "Execute the qryActiveCustomers query and format as CSV"
- "Create a new query called qryRecentOrders with SQL: SELECT * FROM Orders WHERE OrderDate > Date()-30"

## Office Security Settings

⚠️ **Required Setting**: You must enable VBA project access in Office:

1. Open Excel or Access
2. Go to **File** → **Options** → **Trust Center**
3. Click **Trust Center Settings**
4. Select **Macro Settings**
5. Check ✅ **Trust access to the VBA project object model**

See [docs/SECURITY.md](docs/SECURITY.md) for detailed instructions.

## Available Tools

### Excel VBA Tools

| Tool | Description |
|------|-------------|
| `list_open_excel_files` | List currently open Excel workbooks |
| `list_excel_vba_modules` | List all VBA modules in an Excel workbook |
| `read_excel_vba_module` | Read entire module code from Excel |
| `write_excel_vba_module` | Write/replace module code in Excel |
| `create_excel_vba_module` | Create a new module in Excel |
| `delete_excel_vba_module` | Delete a module from Excel |
| `list_excel_vba_procedures` | List procedures in an Excel module |
| `read_excel_vba_procedure` | Read a specific procedure from Excel |
| `write_excel_vba_procedure` | Write/replace a procedure in Excel |
| `export_excel_vba_module` | Export Excel module to file |

### Access VBA Tools

| Tool | Description |
|------|-------------|
| `list_open_access_files` | List currently open Access databases |
| `list_access_vba_modules` | List all VBA modules in an Access database |
| `read_access_vba_module` | Read entire module code from Access |
| `write_access_vba_module` | Write/replace module code in Access |
| `create_access_vba_module` | Create a new module in Access |
| `delete_access_vba_module` | Delete a module from Access |
| `list_access_vba_procedures` | List procedures in an Access module |
| `read_access_vba_procedure` | Read a specific procedure from Access |
| `write_access_vba_procedure` | Write/replace a procedure in Access |
| `export_access_vba_module` | Export Access module to file |

### Access Data Tools

#### Table and Query Operations

| Tool | Description |
|------|-------------|
| `list_access_tables` | List all tables in an Access database |
| `get_access_table_structure` | Get field definitions for a table |
| `get_access_table_data` | Query table data with WHERE clause support |
| `list_access_queries` | List all saved queries in the database |
| `get_access_query_sql` | Get SQL text of a saved query |
| `execute_access_query` | Execute a saved query and return results (supports parameters) |
| `save_access_query` | Create or update a saved query |
| `delete_access_query` | Delete a saved query |

#### Relationship and Index Information

| Tool | Description |
|------|-------------|
| `list_access_relationships` | List all relationships between tables |
| `get_access_table_indexes` | Get all indexes for a specific table |

#### Database Information

| Tool | Description |
|------|-------------|
| `get_access_database_info` | Get summary information (file size, table count, query count, etc.) |
| `list_access_forms` | List all forms in the database |
| `list_access_reports` | List all reports in the database |

#### Data Export

| Tool | Description |
|------|-------------|
| `export_access_table_to_csv` | Export table data to a CSV file |
| `export_access_query_to_csv` | Export query results to a CSV file |

**Important Notes:**
- Excel tools work with `.xlsm`, `.xlsb`, `.xls` files
- Access tools work with `.accdb`, `.mdb` files
- All files must be open in their respective Office applications before using these tools
- Access module names may include prefixes like `Form_MainForm` or `Report_Report1` for code-behind modules

## Breaking Changes in v0.4.0

Excel VBA tool names have been updated to include the `excel` prefix for consistency with Access tools:

| Old Name (v0.3.x) | New Name (v0.4.0+) |
|-------------------|---------------------|
| `list_vba_modules` | `list_excel_vba_modules` |
| `read_vba_module` | `read_excel_vba_module` |
| `write_vba_module` | `write_excel_vba_module` |
| `create_vba_module` | `create_excel_vba_module` |
| `delete_vba_module` | `delete_excel_vba_module` |
| `export_vba_module` | `export_excel_vba_module` |
| `list_vba_procedures` | `list_excel_vba_procedures` |
| `read_vba_procedure` | `read_excel_vba_procedure` |
| `write_vba_procedure` | `write_excel_vba_procedure` |

**Action Required**: If you have existing scripts or workflows using the old tool names, please update them to use the new names.

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
2. **「Browse」ボタンをクリックして対象の Excel/Access ファイルを選択**
3. 「Start」ボタンをクリックして MCP サーバーを実行
4. リアルタイムでログを監視

**注意事項:**
- GUI は VbaMcpServer.exe をレジストリエントリ（インストーラーで設定）または同じディレクトリから自動検出します
- 必要に応じて `appsettings.json` でサーバーパスを上書き可能です
- 選択したターゲットファイルはサーバー起動時に自動的に開かれます

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

Claude Code(CLI ツール)の場合:

**Windows:**
```json
{
  "mcpServers": {
    "vba": {
      "command": "C:\\Program Files\\VBA MCP Server\\VbaMcpServer.exe"
    }
  }
}
```

**macOS/Linux:**
```json
{
  "mcpServers": {
    "vba": {
      "command": "/path/to/VbaMcpServer.exe"
    }
  }
}
```

設定ファイルの場所:
- Windows: `%USERPROFILE%\.claude\settings.json`
- macOS/Linux: `~/.claude/settings.json`

## ⚠️ 重要: バックアップとバージョン管理

**本ツールは自動バックアップ機能を提供しません。** VBA コードの変更は不可逆的な操作です。作業内容の保護は利用者の責任で行ってください：

### 推奨される対策

1. **Git で VBA コードを管理**: Git などのバージョン管理システムで VBA コードを管理する
2. **編集前にファイルをバックアップ**: コード変更前に必ず Excel/Access ファイルのコピーを作成する
3. **Office の自動保存を活用**: OneDrive/SharePoint を使用している場合は、自動バージョン履歴機能を活用する

**VBA コードの変更は永続的であり、本ツールでは元に戻せません。変更前に必ずファイルのバックアップを取ってください。**

## 使用例

設定後、Claude に以下のように依頼できます：

### Excel の例

- 「C:\Projects\MyWorkbook.xlsm の VBA モジュール一覧を表示して」
- 「Module1 のコードを見せて」
- 「SaveData プロシージャにエラーハンドリングを追加して」
- 「DataProcessor という新しいクラスモジュールを作成して」
- 「このコードを事前バインディングを使うようにリファクタリングして」

### Access の例

- 「C:\Projects\MyDatabase.accdb の VBA モジュール一覧を表示して」
- 「Form_MainForm モジュールのコードを見せて」
- 「Form_MainForm の btnSave_Click プロシージャにエラーハンドリングを追加して」
- 「DatabaseConnection という新しいクラスモジュールを Access データベースに作成して」

### Access データの例

- 「C:\Projects\MyDatabase.accdb の全テーブルを一覧表示して」
- 「Customers テーブルの構造を見せて」
- 「Orders テーブルから OrderDate > #2024-01-01# の条件で最初の50件を取得して」
- 「データベース内の全クエリを一覧表示して」
- 「qryMonthlyReport クエリの SQL を見せて」
- 「qryActiveCustomers クエリを実行して CSV 形式で返して」
- 「SELECT * FROM Orders WHERE OrderDate > Date()-30 という SQL で qryRecentOrders という新しいクエリを作成して」

## Office セキュリティ設定

⚠️ **必須設定**: Office で VBA プロジェクトへのアクセスを有効にする必要があります：

1. Excel または Access を開く
2. **ファイル** → **オプション** → **トラストセンター** を選択
3. **トラストセンターの設定** をクリック
4. **マクロの設定** を選択
5. ✅ **VBA プロジェクト オブジェクト モデルへのアクセスを信頼する** にチェック

詳細は [docs/SECURITY.md](docs/SECURITY.md) を参照してください。

## 利用可能なツール

### Excel VBA ツール

| ツール | 説明 |
|--------|------|
| `list_open_excel_files` | 開いている Excel ワークブックを一覧表示 |
| `list_excel_vba_modules` | Excel ワークブック内のすべての VBA モジュールを一覧表示 |
| `read_excel_vba_module` | Excel からモジュール全体のコードを読み取り |
| `write_excel_vba_module` | Excel でモジュールコードを書き込み/置換 |
| `create_excel_vba_module` | Excel で新しいモジュールを作成 |
| `delete_excel_vba_module` | Excel からモジュールを削除 |
| `list_excel_vba_procedures` | Excel モジュール内のプロシージャを一覧表示 |
| `read_excel_vba_procedure` | Excel から特定のプロシージャを読み取り |
| `write_excel_vba_procedure` | Excel でプロシージャを書き込み/置換 |
| `export_excel_vba_module` | Excel モジュールをファイルにエクスポート |

### Access VBA ツール

| ツール | 説明 |
|--------|------|
| `list_open_access_files` | 開いている Access データベースを一覧表示 |
| `list_access_vba_modules` | Access データベース内のすべての VBA モジュールを一覧表示 |
| `read_access_vba_module` | Access からモジュール全体のコードを読み取り |
| `write_access_vba_module` | Access でモジュールコードを書き込み/置換 |
| `create_access_vba_module` | Access で新しいモジュールを作成 |
| `delete_access_vba_module` | Access からモジュールを削除 |
| `list_access_vba_procedures` | Access モジュール内のプロシージャを一覧表示 |
| `read_access_vba_procedure` | Access から特定のプロシージャを読み取り |
| `write_access_vba_procedure` | Access でプロシージャを書き込み/置換 |
| `export_access_vba_module` | Access モジュールをファイルにエクスポート |

### Access データツール

#### テーブル・クエリ操作

| ツール | 説明 |
|--------|------|
| `list_access_tables` | Access データベース内のすべてのテーブルを一覧表示 |
| `get_access_table_structure` | テーブルのフィールド定義を取得 |
| `get_access_table_data` | WHERE句対応でテーブルデータを取得 |
| `list_access_queries` | データベース内のすべての保存済みクエリを一覧表示 |
| `get_access_query_sql` | 保存済みクエリのSQL文を取得 |
| `execute_access_query` | 保存済みクエリを実行して結果を返す（パラメータ対応） |
| `save_access_query` | 保存済みクエリを作成または更新 |
| `delete_access_query` | 保存済みクエリを削除 |

#### リレーションシップ・インデックス情報

| ツール | 説明 |
|--------|------|
| `list_access_relationships` | テーブル間のすべてのリレーションシップを一覧表示 |
| `get_access_table_indexes` | 特定のテーブルのすべてのインデックスを取得 |

#### データベース情報

| ツール | 説明 |
|--------|------|
| `get_access_database_info` | サマリー情報を取得（ファイルサイズ、テーブル数、クエリ数など） |
| `list_access_forms` | データベース内のすべてのフォームを一覧表示 |
| `list_access_reports` | データベース内のすべてのレポートを一覧表示 |

#### データエクスポート

| ツール | 説明 |
|--------|------|
| `export_access_table_to_csv` | テーブルデータをCSVファイルにエクスポート |
| `export_access_query_to_csv` | クエリ結果をCSVファイルにエクスポート |

**重要事項:**
- Excel ツールは `.xlsm`, `.xlsb`, `.xls` ファイルに対応
- Access ツールは `.accdb`, `.mdb` ファイルに対応
- すべてのファイルは各 Office アプリケーションで開いている必要があります
- Access モジュール名には `Form_MainForm` や `Report_Report1` のようなプレフィックスが含まれる場合があります

## v0.4.0 の破壊的変更

Excel VBA ツール名が Access ツールとの一貫性のために `excel` プレフィックスを含むように更新されました：

| 旧名前 (v0.3.x) | 新名前 (v0.4.0+) |
|-----------------|------------------|
| `list_vba_modules` | `list_excel_vba_modules` |
| `read_vba_module` | `read_excel_vba_module` |
| `write_vba_module` | `write_excel_vba_module` |
| `create_vba_module` | `create_excel_vba_module` |
| `delete_vba_module` | `delete_excel_vba_module` |
| `export_vba_module` | `export_excel_vba_module` |
| `list_vba_procedures` | `list_excel_vba_procedures` |
| `read_vba_procedure` | `read_excel_vba_procedure` |
| `write_vba_procedure` | `write_excel_vba_procedure` |

**対応が必要**: 既存のスクリプトやワークフローで旧ツール名を使用している場合は、新しい名前に更新してください。

## ライセンス

MIT License - 詳細は [LICENSE](LICENSE) ファイルを参照してください。
