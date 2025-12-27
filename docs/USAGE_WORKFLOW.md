# VBA MCP Server Usage Workflow / VBA MCP Server 使用ワークフロー

[English](#english) | [日本語](#japanese)

---

<a name="english"></a>

## Understanding Prerequisites

VBA MCP Server uses **COM interop**, so **an Excel or Access application instance must actually be running**.

### Why Do You Need to Open the Application?

1. **Access to COM Object Model**
   - VbaMcpServer connects to a running Excel/Access instance
   - Uses `Marshal.GetActiveObject("Excel.Application")` to get existing instance
   - Accesses VBA through the application, not by directly parsing files

2. **VBA Project Manipulation**
   - Reading and writing VBA code requires the `VBProject` object
   - This is only available from a running workbook/database

3. **Reliability and Compatibility**
   - Works with all versions by using Office's official API
   - No need for detailed knowledge of file formats

## Basic Workflow

### Step 1: Configure Excel

#### 1.1 Trust Center Settings

Allow COM access to VBA projects.

**Excel 2016/2019/365:**
1. Open **File** > **Options**
2. Select **Trust Center**
3. Click **Trust Center Settings** button
4. Select **Macro Settings**
5. Check ✅ **Trust access to the VBA project object model**
6. Click **OK**

**The same settings are required for Access.**

⚠️ **Security Notice**:
- This setting allows external applications to modify VBA code
- Only enable in trusted environments
- Recommended to disable when not in use

#### 1.2 Create/Open Macro-Enabled Workbook

```
Method 1: Open existing file
  - Open an .xlsm (Excel Macro-Enabled Workbook) file

Method 2: Create new
  - Launch Excel
  - File > Save As > Select "Excel Macro-Enabled Workbook (*.xlsm)" as file type
```

#### 1.3 Prepare VBA Project

At minimum, add one module:

```
1. Press Alt + F11 to open VBA Editor
2. Insert > Module to create Module1
3. Write simple code (example):

Sub Test()
    MsgBox "Hello, VBA MCP Server!"
End Sub
```

### Step 2: Start MCP Server

#### Using GUI Version (Recommended)

```bash
# Start GUI manager
.\bin\Debug\VbaMcpServer.GUI.exe

# Or from Visual Studio
dotnet run --project src/VbaMcpServer.GUI
```

**GUI Operation:**
1. When app starts, Server Log tab shows path detection logs
2. Confirm "✓ Found VbaMcpServer.exe" is displayed
3. Click **Start** button
4. Status becomes **Running** (green)

#### Using CLI Version

```bash
# Run server directly
.\bin\Debug\VbaMcpServer.exe
```

**Verification:**
- Console shows "MCP Server started successfully"
- If no errors, MCP protocol is waiting on stdin/stdout

### Step 3: Integration with Claude Desktop

#### Configure Claude Desktop

Edit `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "vba": {
      "command": "C:\\path\\to\\bin\\Debug\\VbaMcpServer.exe"
    }
  }
}
```

**For Installer Version:**
```json
{
  "mcpServers": {
    "vba": {
      "command": "C:\\Program Files\\VBA MCP Server\\VbaMcpServer.exe"
    }
  }
}
```

#### Restart Claude Desktop

Restart Claude Desktop to apply the configuration.

### Step 4: Working with VBA Code

You can ask Claude Desktop questions like:

#### Example 1: List Workbooks

```
User: Show me the open Excel workbooks

Claude: The currently open workbooks are:
- C:\Work\sample.xlsm
```

#### Example 2: List Modules

```
User: List the VBA modules in sample.xlsm

Claude: sample.xlsm has the following modules:
- Module1 (Standard Module)
- ThisWorkbook (Document Module)
- Sheet1 (Document Module)
```

#### Example 3: Read Code

```
User: Show me the code in Module1

Claude: The code in Module1 is:

Sub Test()
    MsgBox "Hello, VBA MCP Server!"
End Sub
```

#### Example 4: Write Code

```
User: Add a new Sub procedure to Module1.
Name it DebugPrint and make it output data to the debug window.

Claude: I've added the DebugPrint procedure to Module1:

Sub DebugPrint(data As Variant)
    Debug.Print data
End Sub
```

**Automatic Backup:**
- Backups are automatically created when writing code
- Saved to: `%USERPROFILE%\.vba-mcp-server\backups\`

## Troubleshooting

### Problem 1: "Excel is not running"

**Symptom:**
```
Error: Excel is not running
```

**Cause:**
- Excel is not running
- Or, no workbook is open

**Solution:**
1. Launch Excel
2. Open .xlsm file
3. Restart MCP Server

### Problem 2: "Access denied to VBA project"

**Symptom:**
```
Error: Access denied to VBA project
```

**Cause:**
- "Trust access to the VBA project object model" is disabled

**Solution:**
1. Excel Options > Trust Center > Trust Center Settings
2. Macro Settings > ✅ Trust access to the VBA project object model
3. Restart Excel

### Problem 3: "Module not found"

**Symptom:**
```
Error: Module 'Module2' not found in workbook
```

**Cause:**
- The specified module name doesn't exist

**Solution:**
1. Press Alt + F11 to open VBA Editor
2. Check module list in Project Explorer
3. Use correct module name

### Problem 4: Version Mismatch Error

**Symptom:**
```
System.IO.FileNotFoundException: Could not load file or assembly 'Microsoft.Extensions.Logging.Abstractions, Version=10.0.0.0'
```

**Cause:**
- Dependency DLL version mismatch

**Solution:**
1. Exit GUI
2. Rebuild solution:
   ```bash
   dotnet build
   ```
3. Restart GUI

## Advanced Usage

### Working with Multiple Workbooks

```
User: Read Module1 from all open workbooks

Claude: (Reads Module1 from each workbook in sequence)
```

### Procedure-Level Operations

```
User: Rewrite only the Test function in Module1

Claude: Updated the Test function.
```

### Checking Backups

Backups are saved to:

```
%USERPROFILE%\.vba-mcp-server\backups\
├── sample.xlsm_Module1_20251226_103045.bas
├── sample.xlsm_Module1_20251226_103112.bas
└── ...
```

**Restore:**
1. Open backup file (.bas)
2. In VBA Editor: "File > Import File"
3. Select backup file

## Best Practices

### 1. Pre-Work Preparation

- [ ] Verify Excel settings (Trust Center)
- [ ] Open target workbook
- [ ] Start MCP Server
- [ ] Restart Claude Desktop

### 2. Safe Operation

- [ ] Make copy of important files beforehand
- [ ] Confirm backup feature is enabled
- [ ] Manual backup before major changes

### 3. Performance

- [ ] Close unnecessary workbooks
- [ ] For many modules, work individually
- [ ] Restart Excel after long sessions

## Summary

VBA MCP Server **accesses VBA code through running Excel/Access instances**. This provides:

✅ **Benefits:**
- Safe operations using Office's official API
- Works with all Office versions
- Safety through automatic backups

⚠️ **Limitations:**
- Must launch application
- Windows only
- Trust Center settings required

Understanding these limitations makes VBA development much more efficient.

---

<a name="japanese"></a>

## 前提条件の理解

VBA MCP Serverは**COM相互運用**を使用しているため、**ExcelまたはAccessのアプリケーションインスタンスが実際に起動している必要があります**。

### なぜアプリケーションを開く必要があるのか

1. **COMオブジェクトモデルへのアクセス**
   - VbaMcpServerは実行中のExcel/Accessインスタンスに接続
   - `Marshal.GetActiveObject("Excel.Application")` で既存インスタンスを取得
   - ファイルを直接解析するのではなく、アプリケーション経由でVBAにアクセス

2. **VBAプロジェクトの操作**
   - VBAコードの読み書きには `VBProject` オブジェクトが必要
   - これは実行中のワークブック/データベースからのみ取得可能

3. **信頼性と互換性**
   - Officeの公式APIを使用するため、すべてのバージョンで動作
   - ファイル形式の詳細な知識が不要

## 基本ワークフロー

### ステップ1: Excelの設定

#### 1.1 トラストセンターの設定

VBAプロジェクトへのCOMアクセスを許可します。

**Excel 2016/2019/365:**
1. **ファイル** > **オプション** を開く
2. **トラストセンター** を選択
3. **トラストセンターの設定** ボタンをクリック
4. **マクロの設定** を選択
5. ✅ **VBAプロジェクト オブジェクト モデルへのアクセスを信頼する** にチェック
6. **OK** をクリック

**Access でも同様の設定が必要です。**

⚠️ **セキュリティ注意事項**:
- この設定により、外部アプリケーションがVBAコードを変更できるようになります
- 信頼できる環境でのみ有効化してください
- 作業終了後は無効化することを推奨します

#### 1.2 マクロ有効ブックの作成/オープン

```
方法1: 既存ファイルを開く
  - .xlsm (Excel Macro-Enabled Workbook) ファイルを開く

方法2: 新規作成
  - Excel を起動
  - ファイル > 名前を付けて保存 > ファイルの種類で「Excel マクロ有効ブック (*.xlsm)」を選択
```

#### 1.3 VBAプロジェクトの準備

最低限、1つのモジュールを追加しておきます:

```
1. Alt + F11 で VBA エディタを開く
2. 挿入 > 標準モジュール で Module1 を作成
3. 簡単なコードを記述（例）:

Sub Test()
    MsgBox "Hello, VBA MCP Server!"
End Sub
```

### ステップ2: MCP Serverの起動

#### GUI版を使用する場合（推奨）

```bash
# GUIマネージャーを起動
.\bin\Debug\VbaMcpServer.GUI.exe

# または Visual Studio から
dotnet run --project src/VbaMcpServer.GUI
```

**GUI操作:**
1. アプリケーションが起動すると、Server Logタブにパス検出ログが表示される
2. 「✓ Found VbaMcpServer.exe」と表示されることを確認
3. **Start** ボタンをクリック
4. Status が **Running** (緑色) になる

#### CLI版を使用する場合

```bash
# サーバーを直接実行
.\bin\Debug\VbaMcpServer.exe
```

**確認ポイント:**
- コンソールに "MCP Server started successfully" と表示される
- エラーがなければ、stdin/stdoutでMCPプロトコルが待機状態

### ステップ3: Claude Desktop との連携

#### Claude Desktop の設定

`%APPDATA%\Claude\claude_desktop_config.json` を編集:

```json
{
  "mcpServers": {
    "vba": {
      "command": "C:\\path\\to\\bin\\Debug\\VbaMcpServer.exe"
    }
  }
}
```

**インストーラー版の場合:**
```json
{
  "mcpServers": {
    "vba": {
      "command": "C:\\Program Files\\VBA MCP Server\\VbaMcpServer.exe"
    }
  }
}
```

#### Claude Desktop の再起動

設定を反映させるため、Claude Desktop を再起動します。

### ステップ4: VBAコードの操作

Claude Desktop で以下のように質問できます:

#### 例1: ワークブックの一覧

```
User: 開いているExcelワークブックを教えて

Claude: 現在開いているワークブックは以下の通りです:
- C:\Work\sample.xlsm
```

#### 例2: モジュールの一覧

```
User: sample.xlsm のVBAモジュールを一覧表示して

Claude: sample.xlsm には以下のモジュールがあります:
- Module1 (標準モジュール)
- ThisWorkbook (ドキュメントモジュール)
- Sheet1 (ドキュメントモジュール)
```

#### 例3: コードの読み取り

```
User: Module1 のコードを表示して

Claude: Module1 のコードは以下の通りです:

Sub Test()
    MsgBox "Hello, VBA MCP Server!"
End Sub
```

#### 例4: コードの書き込み

```
User: Module1 に新しいSubプロシージャを追加して。
データをデバッグウィンドウに出力する DebugPrint という名前で。

Claude: Module1 に DebugPrint プロシージャを追加しました:

Sub DebugPrint(data As Variant)
    Debug.Print data
End Sub
```

**自動バックアップ:**
- コード書き込み時、自動的にバックアップが作成されます
- 保存先: `%USERPROFILE%\.vba-mcp-server\backups\`

## トラブルシューティング

### 問題1: "Excel is not running"

**症状:**
```
Error: Excel is not running
```

**原因:**
- Excelが起動していない
- または、ワークブックが開かれていない

**解決策:**
1. Excelを起動
2. .xlsmファイルを開く
3. MCP Serverを再起動

### 問題2: "Access denied to VBA project"

**症状:**
```
Error: Access denied to VBA project
```

**原因:**
- 「VBAプロジェクト オブジェクト モデルへのアクセスを信頼する」が無効

**解決策:**
1. Excel オプション > トラストセンター > トラストセンターの設定
2. マクロの設定 > ✅ VBAプロジェクト オブジェクト モデルへのアクセスを信頼する
3. Excelを再起動

### 問題3: "Module not found"

**症状:**
```
Error: Module 'Module2' not found in workbook
```

**原因:**
- 指定したモジュール名が存在しない

**解決策:**
1. Alt + F11 でVBAエディタを開く
2. プロジェクトエクスプローラーでモジュール一覧を確認
3. 正しいモジュール名を使用

### 問題4: バージョン不一致エラー

**症状:**
```
System.IO.FileNotFoundException: Could not load file or assembly 'Microsoft.Extensions.Logging.Abstractions, Version=10.0.0.0'
```

**原因:**
- 依存DLLのバージョン不一致

**解決策:**
1. GUIを終了
2. ソリューションをリビルド:
   ```bash
   dotnet build
   ```
3. GUIを再起動

## 高度な使用方法

### 複数ワークブックの操作

```
User: すべての開いているワークブックのModule1を読み取って

Claude: (各ワークブックのModule1を順番に読み取り)
```

### プロシージャ単位の操作

```
User: Module1のTest関数だけを書き換えて

Claude: Test関数を更新しました。
```

### バックアップの確認

バックアップは以下の場所に保存されます:

```
%USERPROFILE%\.vba-mcp-server\backups\
├── sample.xlsm_Module1_20251226_103045.bas
├── sample.xlsm_Module1_20251226_103112.bas
└── ...
```

**リストア:**
1. バックアップファイル(.bas)を開く
2. VBAエディタで「ファイル > ファイルのインポート」
3. バックアップファイルを選択

## ベストプラクティス

### 1. 作業前の準備

- [ ] Excelの設定を確認（トラストセンター）
- [ ] 対象ワークブックを開く
- [ ] MCP Serverを起動
- [ ] Claude Desktopを再起動

### 2. 安全な運用

- [ ] 重要なファイルは事前にコピーを作成
- [ ] バックアップ機能が有効なことを確認
- [ ] 大規模な変更前に手動バックアップ

### 3. パフォーマンス

- [ ] 不要なワークブックは閉じる
- [ ] 大量のモジュールがある場合は個別に操作
- [ ] 長時間の作業後はExcelを再起動

## まとめ

VBA MCP Serverは、**実行中のExcel/Accessインスタンスを介してVBAコードにアクセス**します。これにより:

✅ **利点:**
- Officeの公式APIを使用した安全な操作
- すべてのOfficeバージョンで動作
- 自動バックアップによる安全性

⚠️ **制約:**
- アプリケーションを起動する必要がある
- Windows専用
- トラストセンターの設定が必要

この制約を理解した上で使用すれば、VBA開発が大幅に効率化されます。
