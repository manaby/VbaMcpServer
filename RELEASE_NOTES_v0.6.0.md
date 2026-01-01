# VBA MCP Server v0.6.0 Release Notes

## 🎉 新機能

### GUI State Machine 実装

- **11 状態の State Machine** による厳格な状態管理
  - Idle、Starting、Running、Stopping、Error の各状態を細分化
  - 定義された状態遷移のみを許可し、予期しない動作を防止
  - スレッドセーフな実装（lock による排他制御）

### 完全非同期化

- **すべてのサーバー操作を非同期化**
  - Start/Stop/Restart 操作が非同期実行
  - `Thread.Sleep()` を完全削除し、UI フリーズを解消
  - `async/await` パターンによる応答性の向上

### キャンセル対応

- **CancellationToken による操作キャンセル**
  - 長時間実行される操作を中断可能
  - Stop ボタンで起動処理をキャンセル可能
  - クリーンなシャットダウンを保証

### ファイル監視強化

- **リアルタイムファイル監視**
  - Office アプリケーションでのファイル開閉状態を自動検出（5 秒間隔）
  - ユーザーがファイルを誤って閉じた場合に警告バナーを表示
  - `Running_FileClosedByUser` 状態で適切に通知

### サーバークラッシュ検出

- **自動クラッシュ検出**
  - サーバープロセスの異常終了を即座に検出
  - `Error_ServerCrashed` 状態に遷移し、ユーザーに通知
  - エラー情報をログに記録

### UI 強化

- **プログレスバー追加**

  - Starting/Stopping 状態時にマーキー式プログレスバーを表示
  - 処理状況を視覚的にフィードバック

- **警告バナー**

  - ファイルが閉じられた際にオレンジ色の警告バナーを表示
  - 問題の原因を明確に提示

- **状態表示の色分け**
  - 赤: 停止状態
  - オレンジ: 処理中（Starting/Stopping）
  - 緑: 実行中

### Access コントロール操作ツール追加

- **フォーム/レポートのコントロール操作**
  - `get_access_form_controls` - フォームのコントロール一覧取得（サブフォーム対応）
  - `get_access_form_control_properties` - コントロールプロパティ取得
  - `set_access_form_control_property` - コントロールプロパティ設定
  - `get_access_report_controls` - レポートのコントロール一覧取得
  - `get_access_report_control_properties` - レポートコントロールプロパティ取得
  - `set_access_report_control_property` - レポートコントロールプロパティ設定

## 🐛 バグ修正

### COM 参照リーク修正

- **ComObjectWrapper** による自動 COM オブジェクト解放

  - `IDisposable` パターンによる確実なリソース解放
  - `Marshal.ReleaseComObject` の適切な呼び出し

- **ComServiceBase** 基底クラス導入
  - `ExcelComService` と `AccessComService` の共通実装を統一
  - COM オブジェクトのライフサイクル管理を一元化
  - メモリリークを防止

### エラーハンドリング改善

- カスタム例外クラスによる詳細なエラー情報提供
  - `ControlNotFoundException`
  - `FormNotFoundException`
  - `InvalidPropertyValueException`
  - `PropertyNotFoundException`
  - `PropertyReadOnlyException`
  - `ReportNotFoundException`

## 📦 インストール

### MSI インストーラ（推奨）

1. [VbaMcpServer.msi](https://github.com/manaby/VbaMcpServer/releases/download/v0.6.0/VbaMcpServer.msi) をダウンロード
2. インストーラを実行
3. スタートメニューから「VBA MCP Server Manager」を起動

**インストール内容:**

- VbaMcpServer.exe（CLI ツール）
- VbaMcpServer.GUI.exe（GUI マネージャー）
- README.md（ドキュメント）
- LICENSE（ライセンス）
- スタートメニューショートカット
- デスクトップショートカット

**システム要件:**

- Windows 10/11 (64-bit)
- Microsoft Office 2016 以降（Excel/Access）
- .NET Runtime **不要**（Self-contained ビルド）

### ソースからビルド

```bash
git clone https://github.com/manaby/VbaMcpServer.git
cd VbaMcpServer
dotnet build -c Release
```

## 📝 使用方法

### 1. Office の設定

VBA プロジェクトへのアクセスを許可してください:

```
ファイル → オプション → トラストセンター → トラストセンターの設定
→ マクロの設定 → ☑ VBA プロジェクト オブジェクト モデルへのアクセスを信頼する
```

### 2. GUI マネージャーでサーバーを起動

1. スタートメニューから「VBA MCP Server Manager」を起動
2. 「Browse」ボタンで Excel/Access ファイルを選択
3. 「Start」ボタンでサーバーを起動
4. ログビューアでサーバーの動作を確認

### 3. Claude Desktop で使用

`claude_desktop_config.json` に以下を追加:

```json
{
  "mcpServers": {
    "vba": {
      "command": "C:\\Program Files\\VBA MCP Server\\VbaMcpServer.exe"
    }
  }
}
```

## ⚠️ 重要な注意事項

### バックアップについて

**本ツールは自動バックアップ機能を提供しません。** VBA コードの変更は不可逆的な操作です。

**推奨事項:**

1. Git で VBA コードを管理
2. 編集前に Excel/Access ファイル全体をバックアップ
3. OneDrive/SharePoint の自動バージョン履歴を活用

### ファイル制約

- **ローカルファイルのみ対応** - OneDrive/SharePoint 上のファイルは URL 解決の問題により正しく動作しない可能性があります

## 🔧 技術詳細

### アーキテクチャ

- **フレームワーク**: .NET 8.0, WinForms
- **デザインパターン**: State Machine パターン（11 状態）
- **非同期処理**: 完全 async/await 対応
- **COM 管理**: ComObjectWrapper + ComServiceBase

### State Machine 状態一覧

| 状態                       | 説明                                   |
| -------------------------- | -------------------------------------- |
| `Idle_NoFile`              | ファイル未選択                         |
| `Idle_FileSelected`        | ファイル選択済み（停止中）             |
| `Starting_OpeningFile`     | ファイルを開いている（3-13 秒）        |
| `Starting_WaitingForFile`  | ファイルが開くのを待機中（最大 10 秒） |
| `Starting_LaunchingServer` | MCP サーバー起動中（1 秒）             |
| `Running_FileOpen`         | 実行中（正常状態）                     |
| `Running_FileClosedByUser` | 実行中（ファイル閉じ警告）             |
| `Stopping_ServerShutdown`  | サーバー停止中（0-5 秒）               |
| `Stopping_Cleanup`         | クリーンアップ中（瞬時）               |
| `Error_FileOpenFailed`     | エラー: ファイルオープン失敗           |
| `Error_ServerCrashed`      | エラー: サーバークラッシュ             |

## 📊 統計情報

- **MCP ツール数**: 50+ ツール
  - Excel VBA: 12 ツール
  - Access VBA: 12 ツール
  - Access Data: 17 ツール
  - Access Controls: 6 ツール（新規追加）
- **テストケース数**: 26 テスト
- **コード行数**:
  - MainForm.cs: 803 行
  - GuiState.cs: 188 行
  - 全体: 約 5,000 行

## 📄 ライセンス

MIT License

---

**Full Changelog**: https://github.com/manaby/VbaMcpServer/compare/v0.5.0...v0.6.0
