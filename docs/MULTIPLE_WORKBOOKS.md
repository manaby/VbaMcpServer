# Working with Multiple Excel Workbooks / 複数のExcelワークブック操作

[English](#english) | [日本語](#japanese)

---

<a name="english"></a>

## How Target Identification Works

VBA MCP Server identifies workbooks by **full path (complete file path)**.

### Basic Mechanism

1. **Get Excel Instance**
   ```csharp
   Excel.Application excel = GetActiveObject("Excel.Application");
   ```
   - Connect to currently running Excel application
   - If multiple Excel processes exist, use the first one found

2. **Enumerate Workbooks**
   ```csharp
   foreach (Excel.Workbook wb in excel.Workbooks)
   {
       // Loop through all open workbooks
   }
   ```

3. **Match by Full Path**
   ```csharp
   if (string.Equals(wb.FullName, normalizedPath, StringComparison.OrdinalIgnoreCase))
   {
       return wb; // Return matching workbook
   }
   ```

## Usage Examples

### Example 1: List Open Workbooks

**MCP Tool**: `list_open_excel_files`

**Using in Claude Desktop:**
```
User: Show me the Excel files that are currently open

Claude: (Executes list_open_excel_files tool)
The currently open workbooks are:
- C:\Work\Project1.xlsm
- C:\Work\Project2.xlsm
- D:\Documents\Sample.xlsm
```

**Returned JSON:**
```json
{
  "count": 3,
  "workbooks": [
    "C:\\Work\\Project1.xlsm",
    "C:\\Work\\Project2.xlsm",
    "D:\\Documents\\Sample.xlsm"
  ]
}
```

### Example 2: List Modules in Specific Workbook

**MCP Tool**: `list_vba_modules`

**Parameters:**
- `filePath`: Full path of workbook (required)

**Usage:**
```
User: List the VBA modules in C:\Work\Project1.xlsm

Claude: (Executes list_vba_modules("C:\\Work\\Project1.xlsm"))
Project1.xlsm has the following modules:
- Module1 (Standard Module) - 50 lines
- Module2 (Standard Module) - 120 lines
- ThisWorkbook (Document Module) - 15 lines
```

### Example 3: Read Code from Specific Workbook

**MCP Tool**: `read_vba_module`

**Parameters:**
- `filePath`: Full path of workbook (required)
- `moduleName`: Module name (required)

**Usage:**
```
User: Read Module1 from C:\Work\Project1.xlsm

Claude: (Executes read_vba_module("C:\\Work\\Project1.xlsm", "Module1"))
The code in Module1 is:

Sub Test()
    MsgBox "Hello from Project1"
End Sub
```

### Example 4: Process Multiple Workbooks Sequentially

**Conversation with Claude:**
```
User: Read Module1 from all open workbooks

Claude: First, let me check the open workbooks.
(Executes list_open_excel_files)

3 workbooks are open. I'll read them in sequence.

【Module1 in C:\Work\Project1.xlsm】
(Executes read_vba_module("C:\\Work\\Project1.xlsm", "Module1"))
...

【Module1 in C:\Work\Project2.xlsm】
(Executes read_vba_module("C:\\Work\\Project2.xlsm", "Module1"))
...

【Module1 in D:\Documents\Sample.xlsm】
(Executes read_vba_module("D:\\Documents\\Sample.xlsm", "Module1"))
...
```

## Important Points About Path Specification

### 1. Use Full Paths

❌ **NG: Relative paths or filenames only**
```
list_vba_modules("Project1.xlsm")  // Error: Workbook not found
list_vba_modules("..\\Project1.xlsm")  // Error: Workbook not found
```

✅ **OK: Full path**
```
list_vba_modules("C:\\Work\\Project1.xlsm")  // Works correctly
```

### 2. Case Insensitive

All of the following are treated as the same workbook:
```
C:\Work\Project1.xlsm
c:\work\project1.xlsm
C:\WORK\PROJECT1.XLSM
```

### 3. Path Normalization

Internally normalized with `Path.GetFullPath()`, so the following are also considered identical:
```
C:\Work\Project1.xlsm
C:\Work\..\Work\Project1.xlsm
```

## Error Handling

### When Workbook Is Not Open

**Symptom:**
```
Error: Workbook not found or not open: C:\Work\Project1.xlsm. Please open the file in Excel first.
```

**Solution:**
1. Open target file in Excel
2. Verify file path is correct
3. Use `list_open_excel_files` to check actually open files

### When Workbook Cannot Be Found

**Causes:**
- File path is incorrect
- File is not open
- Opened in different Excel instance (see below)

**Debugging Steps:**
```
1. Execute list_open_excel_files
2. Check returned path list
3. Copy and use exact path
```

## When Multiple Excel Processes Exist

### Behavior

When **multiple Excel processes** are running on Windows:

```
excel.exe (PID: 1234)  ← Connected here
├── Project1.xlsm
└── Project2.xlsm

excel.exe (PID: 5678)  ← Not visible
└── Project3.xlsm
```

**Important:** VBA MCP Server connects only to the first Excel instance found by `GetActiveObject("Excel.Application")`.

### Solutions

#### Method 1: Open All Workbooks in Same Excel Instance (Recommended)

```
1. Launch only one Excel
2. Open all workbooks via "File > Open"
```

This ensures all workbooks are opened in the same process.

#### Method 2: Open Only Needed Workbooks

Keep only the workbooks you want to work with open.

#### Method 3: Restart Excel

Close all Excel instances and reopen only needed files.

### Verification Methods

**Task Manager:**
1. Press Ctrl + Shift + Esc to open Task Manager
2. Select "Details" tab
3. Check number of EXCEL.EXE instances

**PowerShell:**
```powershell
Get-Process excel | Select-Object Id, ProcessName, MainWindowTitle
```

## Best Practices

### 1. Optimized Workflow

```
✅ Recommended:
1. Launch Excel
2. Open all needed workbooks
3. Verify with list_open_excel_files
4. Work using full paths

❌ Not Recommended:
1. Launch multiple Excels separately
2. Try to work with filename only
3. Use relative paths
```

### 2. Dialogue Example with Claude

**Efficient approach:**
```
User: First list the open Excel files,
      then compare Module1 in each.

Claude:
1. I'll check the open files
   (Executes list_open_excel_files)

2. I'll read Module1 from each file
   (Executes read_vba_module for each file)

3. I'll compare the code
   ...
```

### 3. Avoiding Errors

```
✅ Good example:
User: Edit Module1 in C:\Work\Project1.xlsm

✅ Even better:
User: Edit Module1 in the currently open workbook
     (Claude automatically lists and selects)

❌ Bad example:
User: Edit Module1 in Project1.xlsm
     (File path is unclear)
```

## Advanced Usage Examples

### Example 1: Batch Processing

```
User: Add a common module "Common" to all open workbooks.
      Include a DebugPrint function that outputs its argument to Debug.Print.

Claude:
1. Check open workbooks
2. Execute write_vba_module for each workbook
3. Add Common module to all workbooks
```

### Example 2: Code Comparison

```
User: Compare Module1 in Project1.xlsm and Project2.xlsm,
      and tell me the differences

Claude:
1. Read Module1 from Project1.xlsm
2. Read Module1 from Project2.xlsm
3. Compare code and report differences
```

### Example 3: Refactoring

```
User: Rename "oldFunction" to "newFunction" in Module1
      across all workbooks

Claude:
1. List open workbooks
2. Read Module1 from each workbook
3. Replace code
4. Write back to each workbook
```

## Summary

### ✅ What You Can Do

- Work with multiple workbooks simultaneously **within same Excel instance**
- Clearly specify targets with full paths
- Claude can automatically list and select workbooks

### ⚠️ Limitations

- Workbooks opened in different Excel processes are not visible
- Full path required (filename only not supported)
- Workbooks must be opened beforehand

### 💡 Recommendations

1. Open all workbooks in one Excel instance
2. Use `list_open_excel_files` to verify, then use full paths
3. Tell Claude "currently open workbooks" and it will enumerate them automatically

---

<a name="japanese"></a>

## 対象の特定方法

VBA MCP Serverは、**フルパス（ファイルの完全パス）でワークブックを特定**します。

### 基本的な仕組み

1. **Excelインスタンスの取得**
   ```csharp
   Excel.Application excel = GetActiveObject("Excel.Application");
   ```
   - 現在実行中のExcelアプリケーションに接続
   - 複数のExcelプロセスがある場合は、最初に見つかったものを使用

2. **ワークブックの列挙**
   ```csharp
   foreach (Excel.Workbook wb in excel.Workbooks)
   {
       // すべての開いているワークブックをループ
   }
   ```

3. **フルパスでの照合**
   ```csharp
   if (string.Equals(wb.FullName, normalizedPath, StringComparison.OrdinalIgnoreCase))
   {
       return wb; // 一致したワークブックを返す
   }
   ```

## 使用例

### 例1: 開いているワークブックの一覧

**MCPツール**: `list_open_excel_files`

**Claude Desktopでの使用:**
```
User: 今開いているExcelファイルを教えて

Claude: (list_open_excel_files ツールを実行)
現在開いているワークブックは以下の通りです:
- C:\Work\Project1.xlsm
- C:\Work\Project2.xlsm
- D:\Documents\Sample.xlsm
```

**返却されるJSON:**
```json
{
  "count": 3,
  "workbooks": [
    "C:\\Work\\Project1.xlsm",
    "C:\\Work\\Project2.xlsm",
    "D:\\Documents\\Sample.xlsm"
  ]
}
```

### 例2: 特定のワークブックのモジュール一覧

**MCPツール**: `list_vba_modules`

**パラメータ:**
- `filePath`: ワークブックのフルパス（必須）

**使用例:**
```
User: C:\Work\Project1.xlsm のVBAモジュールを一覧表示して

Claude: (list_vba_modules("C:\\Work\\Project1.xlsm") を実行)
Project1.xlsm には以下のモジュールがあります:
- Module1 (標準モジュール) - 50行
- Module2 (標準モジュール) - 120行
- ThisWorkbook (ドキュメントモジュール) - 15行
```

### 例3: 特定のワークブックからコードを読み取る

**MCPツール**: `read_vba_module`

**パラメータ:**
- `filePath`: ワークブックのフルパス（必須）
- `moduleName`: モジュール名（必須）

**使用例:**
```
User: C:\Work\Project1.xlsm の Module1 を読み取って

Claude: (read_vba_module("C:\\Work\\Project1.xlsm", "Module1") を実行)
Module1 のコードは以下の通りです:

Sub Test()
    MsgBox "Hello from Project1"
End Sub
```

### 例4: 複数のワークブックを順番に処理

**Claudeとの会話例:**
```
User: 開いているすべてのワークブックのModule1を読み取って

Claude: まず開いているワークブックを確認します。
(list_open_excel_files を実行)

3つのワークブックが開いています。順番に読み取ります。

【C:\Work\Project1.xlsm の Module1】
(read_vba_module("C:\\Work\\Project1.xlsm", "Module1") を実行)
...

【C:\Work\Project2.xlsm の Module1】
(read_vba_module("C:\\Work\\Project2.xlsm", "Module1") を実行)
...

【D:\Documents\Sample.xlsm の Module1】
(read_vba_module("D:\\Documents\\Sample.xlsm", "Module1") を実行)
...
```

## パス指定の重要なポイント

### 1. フルパスを使用

❌ **NG: 相対パスやファイル名のみ**
```
list_vba_modules("Project1.xlsm")  // エラー: ワークブックが見つからない
list_vba_modules("..\\Project1.xlsm")  // エラー: ワークブックが見つからない
```

✅ **OK: フルパス**
```
list_vba_modules("C:\\Work\\Project1.xlsm")  // 正常に動作
```

### 2. 大文字・小文字は区別しない

以下はすべて同じワークブックとして扱われます:
```
C:\Work\Project1.xlsm
c:\work\project1.xlsm
C:\WORK\PROJECT1.XLSM
```

### 3. パスの正規化

内部的に `Path.GetFullPath()` で正規化されるため、以下も同じと判定されます:
```
C:\Work\Project1.xlsm
C:\Work\..\Work\Project1.xlsm
```

## エラーハンドリング

### ワークブックが開かれていない場合

**症状:**
```
Error: Workbook not found or not open: C:\Work\Project1.xlsm. Please open the file in Excel first.
```

**解決策:**
1. Excelで対象ファイルを開く
2. ファイルパスが正しいか確認
3. `list_open_excel_files` で実際に開いているファイルを確認

### ワークブックが見つからない場合

**原因:**
- ファイルパスが間違っている
- ファイルが開かれていない
- 別のExcelインスタンスで開かれている（後述）

**デバッグ手順:**
```
1. list_open_excel_files を実行
2. 返されたパスリストを確認
3. 正確なパスをコピーして使用
```

## 複数のExcelプロセスがある場合

### 動作

Windows上で**複数のExcelプロセス**が起動している場合:

```
excel.exe (PID: 1234)  ← ここに接続される
├── Project1.xlsm
└── Project2.xlsm

excel.exe (PID: 5678)  ← こちらは見えない
└── Project3.xlsm
```

**重要:** VBA MCP Serverは、`GetActiveObject("Excel.Application")` で最初に見つかったExcelインスタンスにのみ接続します。

### 対処方法

#### 方法1: すべてのワークブックを同じExcelインスタンスで開く（推奨）

```
1. Excelを1つだけ起動
2. すべてのワークブックを「ファイル > 開く」で開く
```

これにより、すべてのワークブックが同じプロセスで開かれます。

#### 方法2: 必要なワークブックだけを開く

操作したいワークブックだけを開いた状態にします。

#### 方法3: Excelを再起動

すべてのExcelインスタンスを閉じて、必要なファイルだけを開き直します。

### 確認方法

**タスクマネージャー:**
1. Ctrl + Shift + Esc でタスクマネージャーを開く
2. 「詳細」タブを選択
3. EXCEL.EXE の個数を確認

**PowerShell:**
```powershell
Get-Process excel | Select-Object Id, ProcessName, MainWindowTitle
```

## ベストプラクティス

### 1. ワークフローの最適化

```
✅ 推奨:
1. Excel を起動
2. すべての必要なワークブックを開く
3. list_open_excel_files で確認
4. フルパスを使って操作

❌ 非推奨:
1. 複数のExcelを別々に起動
2. ファイル名だけで操作しようとする
3. 相対パスを使う
```

### 2. Claudeとの対話例

**効率的な方法:**
```
User: まず開いているExcelファイルをリストアップして、
      それぞれのModule1を比較してください。

Claude:
1. 開いているファイルを確認します
   (list_open_excel_files を実行)

2. 各ファイルのModule1を読み取ります
   (ファイルごとに read_vba_module を実行)

3. コードを比較します
   ...
```

### 3. エラー回避

```
✅ 良い例:
User: C:\Work\Project1.xlsm のModule1を編集して

✅ さらに良い例:
User: 今開いているワークブックのModule1を編集して
     (Claudeが自動的にリストアップして選択)

❌ 悪い例:
User: Project1.xlsm のModule1を編集して
     (ファイルパスが不明確)
```

## 高度な使用例

### 例1: バッチ処理

```
User: 開いているすべてのワークブックに共通のモジュール "Common" を追加して。
      内容は、DebugPrint という関数で、引数をDebug.Printに出力するもの。

Claude:
1. 開いているワークブックを確認
2. 各ワークブックに対して write_vba_module を実行
3. すべてのワークブックに Common モジュールを追加
```

### 例2: コードの比較

```
User: Project1.xlsm と Project2.xlsm の Module1 を比較して、
      違いを教えて

Claude:
1. Project1.xlsm の Module1 を読み取り
2. Project2.xlsm の Module1 を読み取り
3. コードを比較して差分をレポート
```

### 例3: リファクタリング

```
User: すべてのワークブックのModule1にある "oldFunction" を
      "newFunction" にリネームして

Claude:
1. 開いているワークブックをリストアップ
2. 各ワークブックのModule1を読み取り
3. コードを置換
4. 各ワークブックに書き戻し
```

## まとめ

### ✅ できること

- 複数のワークブックを**同じExcelインスタンス内で**同時に操作
- フルパスで明確に対象を指定
- Claudeが自動的にワークブックをリストアップして選択

### ⚠️ 制約

- 異なるExcelプロセスで開かれたワークブックは見えない
- フルパスが必須（ファイル名のみは不可）
- ワークブックは事前に開いておく必要がある

### 💡 推奨事項

1. すべてのワークブックを1つのExcelインスタンスで開く
2. `list_open_excel_files` で確認してからフルパスを使用
3. Claudeに「開いているワークブック」と伝えれば自動的に列挙してくれる
