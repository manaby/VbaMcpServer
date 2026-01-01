# VBA MCP Server 要件定義書

## 1. プロジェクト概要

### 1.1 目的
Excel/Access の VBA コードを MCP (Model Context Protocol) 経由で読み書き可能にし、Claude Desktop や Cursor などの AI コーディング環境から VBA の Vibe コーディングを実現する。

### 1.2 背景・課題
- VBA は Office 内蔵 IDE（Visual Basic Editor）からしか編集できない
- AI を利用したコーディングには既存コードのコピペが必要
- 全体の文脈を AI に理解させることが困難
- 他のプログラミング言語と比較して AI 支援の恩恵を受けにくい

### 1.3 解決策
MCP サーバーを介して Excel/Access の VBA プロジェクトに直接アクセスし、AI が VBA コードを読み書きできるようにする。

---

## 2. システム要件

### 2.1 実行環境

| 項目 | 要件 |
|------|------|
| OS | Windows 10/11 (64-bit) |
| .NET | .NET 8.0 Runtime |
| Office | Microsoft Office 2016 以降（Excel/Access） |
| メモリ | 4GB 以上推奨 |

### 2.2 開発環境

| 項目 | バージョン |
|------|------------|
| IDE | Visual Studio 2022 Community 以降 |
| SDK | .NET 8.0 SDK |
| 言語 | C# 12 |

### 2.3 依存パッケージ

| パッケージ | バージョン | 用途 |
|-----------|------------|------|
| ModelContextProtocol | 0.5.0-preview.1 | MCP SDK |
| Microsoft.Extensions.Hosting | 8.0.0 | ホスティング基盤 |
| Microsoft.Extensions.Logging | 8.0.0 | ログ出力 |
| Microsoft.Extensions.Logging.Abstractions | 10.0.0 | ログ抽象化 |
| Microsoft.Extensions.Logging.Console | 8.0.0 | コンソールログ |
| Serilog.Extensions.Hosting | 8.0.0 | Serilogホスティング |
| Serilog.Sinks.Console | 5.0.0 | Serilogコンソール出力 |
| Serilog.Sinks.File | 5.0.0 | Serilogファイル出力 |

---

## 3. 機能要件

### 3.1 Excel VBA 操作機能

#### 3.1.1 ワークブック管理
| 機能 | 説明 | 優先度 |
|------|------|--------|
| ワークブック一覧取得 | 開いている Excel ワークブックの一覧を取得 | 必須 |

#### 3.1.2 モジュール管理
| 機能 | 説明 | 優先度 |
|------|------|--------|
| モジュール一覧取得 | VBA プロジェクト内のモジュール一覧を取得 | 必須 |
| モジュール読み取り | モジュールの VBA コード全体を取得 | 必須 |
| モジュール書き込み | モジュールの VBA コードを置換 | 必須 |
| モジュール作成 | 新規モジュールを作成 | 必須 |
| モジュール削除 | 既存モジュールを削除 | 必須 |
| モジュールエクスポート | モジュールをファイルに出力 | 任意 |

#### 3.1.3 対応モジュールタイプ
| タイプ | 読み取り | 書き込み | 作成 | 削除 | 備考 |
|--------|:--------:|:--------:|:----:|:----:|------|
| 標準モジュール (.bas) | ✅ | ✅ | ✅ | ✅ | 完全対応 |
| クラスモジュール (.cls) | ✅ | ✅ | ✅ | ✅ | 完全対応 |
| ユーザーフォーム (.frm) | ✅ | ✅ | ✅ | ✅ | コードのみ、デザイン不可 |
| ドキュメントモジュール | ✅ | ✅ | ❌ | ❌ | ThisWorkbook, Sheet 等 |

### 3.2 Access VBA 操作機能 ✅ 完了

| 機能 | 説明 | 優先度 | 状態 |
|------|------|--------|------|
| データベース一覧取得 | 開いている Access データベースの一覧を取得 | 必須 | ✅ 実装済み |
| モジュール操作 | Excel と同等の機能（読み取り、書き込み、作成、削除） | 必須 | ✅ 実装済み |
| フォーム/レポートのコードビハインド | フォーム・レポートの VBA コード操作 | 必須 | ✅ 実装済み |
| プロシージャ単位の操作 | 個別プロシージャの読み書き | 任意 | ✅ 実装済み |
| モジュールエクスポート | モジュールをファイルに出力 | 任意 | ✅ 実装済み |

#### 3.2.1 対応モジュールタイプ
| タイプ | 読み取り | 書き込み | 作成 | 削除 | 備考 |
|--------|:--------:|:--------:|:----:|:----:|------|
| 標準モジュール (.bas) | ✅ | ✅ | ✅ | ✅ | 完全対応 |
| クラスモジュール (.cls) | ✅ | ✅ | ✅ | ✅ | 完全対応 |
| ユーザーフォーム (.frm) | ✅ | ✅ | ✅ | ✅ | コードのみ、デザイン不可 |
| フォーム/レポート | ✅ | ✅ | ❌ | ❌ | コードビハインドのみ |

### 3.3 Access データ操作機能（テーブル・クエリ）

#### 3.3.1 テーブル操作（読み取り専用）
| 機能 | 説明 | 優先度 |
|------|------|--------|
| テーブル一覧取得 | データベース内のテーブル一覧を取得 | 必須 |
| テーブル構造取得 | フィールド定義（名前、型、サイズ、制約等）を取得 | 必須 |
| テーブルデータ取得 | WHERE条件とページネーション対応のSELECT | 必須 |

#### 3.3.2 クエリ操作
| 機能 | 説明 | 優先度 |
|------|------|--------|
| クエリ一覧取得 | 保存済みクエリの一覧を取得 | 必須 |
| クエリSQL取得 | クエリのSQL文を取得 | 必須 |
| クエリ実行 | 保存済みクエリを実行して結果を返す | 必須 |
| クエリ編集 | SQL文の変更、新規作成、削除 | 必須 |

#### 3.3.3 データ形式
| 形式 | 対応 | 備考 |
|------|:----:|------|
| JSON | ✅ | デフォルト、構造化データに適する |
| CSV | ✅ | パラメータ指定、データエクスポートに適する |

#### 3.3.4 リレーションシップ・インデックス情報
| 機能 | 説明 | 優先度 |
|------|------|--------|
| リレーションシップ一覧 | テーブル間のリレーションシップ情報を取得 | 推奨 |
| インデックス情報取得 | テーブルのインデックス一覧を取得 | 推奨 |

#### 3.3.5 データベース統計情報
| 機能 | 説明 | 優先度 |
|------|------|--------|
| データベース情報取得 | ファイルサイズ、オブジェクト数などのサマリー | 推奨 |
| フォーム一覧取得 | データベース内のフォーム一覧 | 推奨 |
| レポート一覧取得 | データベース内のレポート一覧 | 推奨 |

#### 3.3.6 データエクスポート
| 機能 | 説明 | 優先度 |
|------|------|--------|
| テーブルCSVエクスポート | テーブルデータをCSVファイルに出力 | 推奨 |
| クエリCSVエクスポート | クエリ結果をCSVファイルに出力 | 推奨 |

#### 3.3.7 パラメータクエリ対応
| 機能 | 説明 | 優先度 |
|------|------|--------|
| パラメータクエリ実行 | クエリパラメータを辞書形式で受け取り実行 | 必須 |

### 3.4 バージョン管理とバックアップ（ユーザー責任）

| 推奨事項 | 説明 |
|----------|------|
| Git によるバージョン管理 | VBA コードを Git で管理することを強く推奨 |
| 事前のファイルコピー | コード変更前に、対象ファイル全体のバックアップコピーを作成 |
| Office の自動保存 | OneDrive/SharePoint 使用時は自動バージョン履歴を活用 |

**重要**: 本ツールは自動バックアップ機能を提供しません。VBA コードの変更は不可逆的な操作となるため、必ず事前にファイル全体のバックアップを取ってください。

---

## 4. MCP ツール定義

### 4.1 ツール一覧

#### Excel VBA ツール

| ツール名 | 説明 |
|----------|------|
| `list_open_excel_files` | 開いている Excel ワークブック一覧を取得 |
| `list_excel_vba_modules` | 指定ワークブックの VBA モジュール一覧を取得 |
| `read_excel_vba_module` | モジュールの VBA コードを読み取り |
| `write_excel_vba_module` | モジュールに VBA コードを書き込み |
| `create_excel_vba_module` | 新規 VBA モジュールを作成 |
| `delete_excel_vba_module` | VBA モジュールを削除 |
| `list_excel_vba_procedures` | モジュール内のプロシージャ一覧を取得 |
| `read_excel_vba_procedure` | 特定プロシージャを読み取り |
| `write_excel_vba_procedure` | プロシージャを書き込み/置換（upsert動作） |
| `add_excel_vba_procedure` | 新規プロシージャを追加（既存時はエラー） |
| `delete_excel_vba_procedure` | プロシージャを削除 |
| `export_excel_vba_module` | モジュールをファイルにエクスポート |

#### Access VBA ツール（12ツール）

**モジュール操作（7ツール）:**

| ツール名 | 説明 |
|----------|------|
| `list_open_access_files` | 開いている Access データベース一覧を取得 |
| `list_access_vba_modules` | 指定データベースの VBA モジュール一覧を取得 |
| `read_access_vba_module` | Access モジュールの VBA コードを読み取り |
| `write_access_vba_module` | Access モジュールに VBA コードを書き込み |
| `create_access_vba_module` | 新規 VBA モジュールを Access に作成 |
| `delete_access_vba_module` | Access の VBA モジュールを削除 |
| `export_access_vba_module` | Access モジュールをファイルにエクスポート |

**プロシージャ操作（5ツール）:**

| ツール名 | 説明 |
|----------|------|
| `list_access_vba_procedures` | Access モジュール内のプロシージャ一覧を取得 |
| `read_access_vba_procedure` | Access の特定プロシージャを読み取り |
| `write_access_vba_procedure` | Access のプロシージャを書き込み/置換（upsert動作） |
| `add_access_vba_procedure` | 新規プロシージャを追加（既存時はエラー） |
| `delete_access_vba_procedure` | プロシージャを削除 |

#### Access データツール

**テーブル・クエリ操作:**

| ツール名 | 説明 |
|----------|------|
| `list_access_tables` | Access データベース内のテーブル一覧を取得 |
| `get_access_table_structure` | 指定テーブルのフィールド定義を取得 |
| `get_access_table_data` | テーブルからデータを取得（WHERE句、ページネーション対応） |
| `list_access_queries` | Access データベース内の保存済みクエリ一覧を取得 |
| `get_access_query_sql` | 指定クエリのSQL文を取得 |
| `execute_access_query` | 保存済みクエリを実行（パラメータ、ページネーション対応） |
| `save_access_query` | 保存済みクエリを作成または更新 |
| `delete_access_query` | 保存済みクエリを削除 |

**リレーションシップ・インデックス:**

| ツール名 | 説明 |
|----------|------|
| `list_access_relationships` | テーブル間のリレーションシップ一覧を取得 |
| `get_access_table_indexes` | 指定テーブルのインデックス一覧を取得 |

**データベース情報:**

| ツール名 | 説明 |
|----------|------|
| `get_access_database_info` | データベースのサマリー情報（サイズ、オブジェクト数等）を取得 |
| `list_access_forms` | データベース内のフォーム一覧を取得 |
| `list_access_reports` | データベース内のレポート一覧を取得 |

**データエクスポート:**

| ツール名 | 説明 |
|----------|------|
| `export_access_table_to_csv` | テーブルデータをCSVファイルにエクスポート |
| `export_access_query_to_csv` | クエリ結果をCSVファイルにエクスポート |

### 4.2 ツール詳細

#### ListOpenExcelFiles
```
入力: なし
出力: {
  count: number,
  workbooks: string[]  // ファイルパスの配列
}
```

#### ListVbaModules
```
入力: {
  filePath: string  // ワークブックのフルパス
}
出力: {
  file: string,
  moduleCount: number,
  modules: [{
    name: string,
    type: string,
    lineCount: number,
    procedureCount: number
  }]
}
```

#### ReadVbaModule
```
入力: {
  filePath: string,
  moduleName: string
}
出力: {
  file: string,
  module: string,
  lineCount: number,
  code: string
}
```

#### WriteVbaModule
```
入力: {
  filePath: string,
  moduleName: string,
  code: string
}
出力: {
  success: boolean,
  file: string,
  module: string,
  linesWritten: number
}
```

#### CreateVbaModule
```
入力: {
  filePath: string,
  moduleName: string,
  moduleType: "standard" | "class" | "userform"  // default: "standard"
}
出力: {
  success: boolean,
  file: string,
  module: string,
  type: string
}
```

#### DeleteVbaModule
```
入力: {
  filePath: string,
  moduleName: string
}
出力: {
  success: boolean,
  file: string,
  module: string,
  deleted: boolean
}
```

#### ExportVbaModule
```
入力: {
  filePath: string,
  moduleName: string,
  outputPath: string
}
出力: {
  success: boolean,
  file: string,
  module: string,
  exportedTo: string
}
```

---

## 5. 非機能要件

### 5.1 セキュリティ

| 要件 | 説明 |
|------|------|
| VBA プロジェクトアクセス | Office の「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」設定が必須 |
| ローカル実行 | サーバーはローカルマシンでのみ実行（リモートアクセス不可） |
| ユーザー責任のバックアップ | 破壊的操作の前にユーザー自身がファイルのバックアップを取ること |

### 5.2 パフォーマンス

| 要件 | 目標値 |
|------|--------|
| モジュール読み取り | 1秒以内 |
| モジュール書き込み | 1秒以内 |
| ワークブック一覧取得 | 500ms 以内 |

### 5.3 配布形式

| 形式 | 説明 |
|------|------|
| MSI インストーラ（推奨） | WiX Toolset で作成、GUI管理アプリ含む |
| Self-contained exe | .NET ランタイム同梱の単一実行ファイル（CLI版） |
| サイズ | MSI: 約 80-100MB、CLI exe: 約 60-80MB |
| 対応アーキテクチャ | win-x64 |
| GUI管理ツール | VBA MCP Server Manager（WinForms） |

---

## 6. 制約事項

### 6.1 技術的制約
1. **Windows 専用**: COM 依存のため Windows でのみ動作
2. **Office 起動必須**: 対象ファイルが Office アプリケーションで開いている必要あり
3. **パスワード保護非対応**: パスワード保護された VBA プロジェクトにはアクセス不可
4. **フォームデザイン不可**: ユーザーフォームのコントロール配置は編集不可（コードのみ）

### 6.2 MCP SDK 制約
- SDK はプレビュー版（0.5.0-preview.1）のため、API が変更される可能性あり

---

## 7. 前提条件

### 7.1 ユーザー側の設定（必須）

Excel/Access で以下の設定を有効にする必要があります：

```
ファイル → オプション → トラストセンター → トラストセンターの設定
→ マクロの設定 → ☑ VBA プロジェクト オブジェクト モデルへのアクセスを信頼する
```

### 7.2 対応ファイル形式

| アプリケーション | 対応形式 |
|------------------|----------|
| Excel | .xlsm, .xlsb, .xls, .xltm |
| Access | .accdb, .mdb（将来対応） |

---

## 8. MCP クライアント連携

### 8.1 Claude Desktop 設定例

`claude_desktop_config.json`:
```json
{
  "mcpServers": {
    "vba": {
      "command": "C:\\path\\to\\VbaMcpServer.exe"
    }
  }
}
```

### 8.2 対応 MCP クライアント
- Claude Desktop
- Cursor
- その他 MCP 対応エディタ/ツール

---

## 9. 開発ロードマップ

### Phase 1: Excel 基本機能（現在）
- [x] Excel COM 接続
- [x] モジュール読み書き
- [x] MCP サーバー実装

### Phase 2: 安定化・テスト ✅ 完了
- [x] 単体テスト追加 (26テスト実装)
- [x] エラーハンドリング強化 (カスタム例外クラス5種)
- [x] ログ機能強化 (Serilog導入、サーバー/VBA編集ログ分離)
- [x] プロジェクト構造再編成 (Core/CLI/GUI/Tests分離)
- [x] GUI管理アプリケーション (WinForms)
- [x] MSIインストーラ (WiX Toolset)

### Phase 3: Access 対応 ✅ 完了 (commit 730a07e - 2025-12-28)
- [x] Access COM 接続 (AccessComService 実装)
- [x] フォーム/レポートのコードビハインド対応
- [x] モジュール CRUD 操作
- [x] プロシージャ単位の操作
- [x] AccessVbaTools 10ツール実装

### Phase 4: 拡張機能 ✅ 完了 (commit xxxxx - 2025-12-29)
- [x] プロシージャ単位の操作（upsert動作、add/delete対応）
- [x] XMLエスケープ防御処理
- [x] 改行コード正規化
- [ ] 部分的なコード編集（行単位）
- [ ] 参照設定の管理
- [ ] VBA プロジェクトのインポート

### Phase 5: GUI State Machine実装 ✅ 完了 (commit eee76e8 - 2025-12-31)
- [x] 11状態のState Machine実装（GuiState.cs）
- [x] 状態遷移の厳格化（許可された遷移のみ実行）
- [x] すべての操作の非同期化（Start/Stop/Restart）
- [x] CancellationTokenによるキャンセル対応
- [x] ファイル監視とState連携（Running_FileClosedByUser状態検出）
- [x] サーバークラッシュ検出（Error_ServerCrashed状態）
- [x] UI強化：プログレスバー（Starting/Stopping状態で表示）
- [x] UI強化：警告バナー（ファイル閉じ検出時に表示）
- [x] Thread.Sleep削除によるUIフリーズ解消
- [x] COM参照リーク修正（ComObjectWrapper、ComServiceBase）

#### GUI管理アプリケーション詳細仕様

**アーキテクチャ:**
- **フレームワーク**: WinForms (.NET 8.0)
- **デザインパターン**: State Machine パターン（11状態）
- **非同期処理**: 完全async/await対応
- **キャンセル**: CancellationToken による操作キャンセル対応

**UI構成（3グループボックス）:**

1. **Target File グループボックス:**
   - ファイルパステキストボックス（読み取り専用）
   - Browse ボタン - Excel/Accessファイル選択（OpenFileDialog使用）
   - Clear ボタン - 選択ファイルクリア
   - ファイルステータスラベル - Officeアプリでの開閉状態表示
   - 警告バナーパネル（オレンジ背景） - ファイル閉じ検出時に表示

2. **Server Control グループボックス:**
   - ステータスラベル - サーバー状態表示（色分け: 赤=停止、橙=処理中、緑=実行中）
   - プロセスIDラベル - サーバープロセスID表示
   - Start ボタン - MCPサーバー起動
   - Stop ボタン - サーバー停止
   - Restart ボタン - サーバー再起動
   - Force Stop ボタン - 強制停止（タイムアウト時表示、将来実装）
   - プログレスバー - Starting/Stopping状態時に表示（マーキー式）

3. **Log Viewer グループボックス:**
   - タブコントロール（2タブ）:
     - **Server Log タブ** - MCPサーバー出力（黒背景、緑文字）
     - **VBA Edit Log タブ** - VBA編集履歴
   - Clear ボタン - 現在のタブのログクリア
   - Save ボタン - ログをテキストファイルに保存（SaveFileDialog使用）

**State Machine（11状態の遷移表）:**

| 状態 | 説明 | 許可される遷移先 |
|------|------|------------------|
| `Idle_NoFile` | ファイル未選択 | `Idle_FileSelected` |
| `Idle_FileSelected` | ファイル選択済み（サーバー停止中） | `Idle_NoFile`, `Starting_OpeningFile` |
| `Starting_OpeningFile` | ファイルを開いている（3-13秒） | `Starting_WaitingForFile`, `Error_FileOpenFailed`, `Stopping_Cleanup` |
| `Starting_WaitingForFile` | ファイルが開くのを待機中（最大10秒） | `Starting_LaunchingServer`, `Error_FileOpenFailed`, `Stopping_Cleanup` |
| `Starting_LaunchingServer` | MCPサーバー起動中（1秒） | `Running_FileOpen`, `Stopping_ServerShutdown` |
| `Running_FileOpen` | 実行中（ファイル開いている正常状態） | `Running_FileClosedByUser`, `Stopping_ServerShutdown`, `Error_ServerCrashed` |
| `Running_FileClosedByUser` | 実行中（ユーザーがファイルを手動で閉じた） | `Running_FileOpen`, `Stopping_ServerShutdown`, `Error_ServerCrashed` |
| `Stopping_ServerShutdown` | サーバープロセス停止中（0-5秒） | `Stopping_Cleanup` |
| `Stopping_Cleanup` | クリーンアップ処理中（瞬時） | `Idle_FileSelected`, `Starting_OpeningFile` |
| `Error_FileOpenFailed` | エラー: ファイルオープン失敗 | `Idle_NoFile`, `Idle_FileSelected`, `Starting_OpeningFile` |
| `Error_ServerCrashed` | エラー: サーバープロセスクラッシュ | `Idle_NoFile`, `Idle_FileSelected`, `Starting_OpeningFile` |

**状態遷移ルール:**
- 定義された遷移のみ許可（InvalidOperationException発生）
- スレッドセーフ（lockによる排他制御）
- 状態変化イベント発火（UIスレッドで実行）

**主要機能:**

1. **非同期処理:**
   - すべてのサーバー操作（Start/Stop/Restart）を非同期実行
   - Task.Run() によるバックグラウンドスレッド実行
   - Thread.Sleep() 削除によるUIフリーズ解消

2. **ファイル監視:**
   - FileOpenerService によるリアルタイム監視（5秒間隔）
   - Officeアプリケーションでのファイル開閉状態検出
   - プロセスID取得とステータス表示

3. **ログ管理:**
   - LogViewerService による2種類のログ監視
   - リアルタイムログ追加とファイル保存機能
   - タブ式UI（Server Log / VBA Edit Log）

4. **COM参照リーク対策:**
   - ComObjectWrapper による自動解放
   - ComServiceBase 基底クラスによる統一的なCOM管理
   - using パターンによるリソース解放保証

**技術的実装:**
- **ファイル**: `MainForm.cs` (803行)、`MainForm.Designer.cs`、`GuiState.cs` (188行)
- **サービス**: `McpServerHostService.cs`、`LogViewerService.cs`、`FileOpenerService.cs`
- **実装完了コミット**: eee76e8（Phase 2完了: COM参照リーク修正とコード品質向上）

---

## 10. ライセンス

- **プロジェクト**: MIT License
- **.NET Runtime**: MIT License
- **MCP SDK**: MIT License

---

## 改訂履歴

| バージョン | 日付 | 変更内容 |
|------------|------|----------|
| 0.6.0 | 2025-12-31 | Phase 5完了: GUI State Machine実装（11状態、非同期化、UIフリーズ解消、COM参照リーク修正） |
| 0.5.0 | 2025-12-29 | Phase 4完了: プロシージャ操作強化（upsert動作、add/delete対応）、XMLエスケープ防御、改行正規化 |
| 0.4.0 | 2025-12-29 | 破壊的変更: Excelツール名に`excel`プレフィックス追加、Access VBA対応ドキュメント追加 |
| 0.3.0 | 2025-12-28 | Phase 3完了: Access VBA対応実装（AccessComService, AccessVbaTools 10ツール） |
| 0.2.0 | 2025-12-26 | Phase 2完了: GUI管理アプリ追加、単体テスト追加、ログ強化、MSIインストーラ対応 |
| 0.1.0 | 2024-12-26 | 初版作成 |
