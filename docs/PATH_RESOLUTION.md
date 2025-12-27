# VbaMcpServer.exe Path Resolution Strategy / VbaMcpServer.exe パス解決方法

[English](#english) | [日本語](#japanese)

---

<a name="english"></a>

## Overview

This document explains how the VBA MCP Server Manager GUI locates the VbaMcpServer.exe executable in both production (installed) and development environments.

## Design Goals

1. **Zero configuration for standard installations**: Works immediately after MSI installation
2. **Development-friendly**: Automatically detects development builds without configuration
3. **Flexible**: Allows user override for custom scenarios
4. **Robust**: Falls back gracefully if preferred methods fail
5. **Transparent**: Logs all path resolution attempts for debugging

## Resolution Priority Order

The GUI searches for VbaMcpServer.exe in the following order (first match wins):

**Note**: The search process uses only 3 stages. Development build auto-detection is not implemented.

### Priority 1: User Configuration (appsettings.json)

**File**: `appsettings.json` (same directory as VbaMcpServer.GUI.exe)

```json
{
  "VbaMcpServer": {
    "ServerExePath": "C:\\Custom\\Path\\VbaMcpServer.exe"
  }
}
```

**When to use**:
- Custom installation directory
- Network share location
- Non-standard deployment
- Development environment with unusual structure

**Example scenarios**:
```json
// Custom installation
{ "VbaMcpServer": { "ServerExePath": "D:\\Tools\\VbaMcpServer.exe" } }

// Network share
{ "VbaMcpServer": { "ServerExePath": "\\\\server\\tools\\VbaMcpServer.exe" } }

// Relative path (not recommended)
{ "VbaMcpServer": { "ServerExePath": "..\\..\\bin\\VbaMcpServer.exe" } }
```

### Priority 2: Windows Registry

**Registry Key**: `HKEY_CURRENT_USER\Software\VbaMcpServer`
**Value Name**: `ServerExePath`
**Value Type**: `REG_SZ`
**Example Value**: `C:\Program Files\VBA MCP Server\VbaMcpServer.exe`

**Set by**: MSI installer during installation
**When to use**: Standard installation via MSI

**Manual registry setup** (if needed):
```powershell
# PowerShell
New-Item -Path "HKCU:\Software\VbaMcpServer" -Force
Set-ItemProperty -Path "HKCU:\Software\VbaMcpServer" -Name "ServerExePath" -Value "C:\Program Files\VBA MCP Server\VbaMcpServer.exe"

# Command Prompt
reg add "HKCU\Software\VbaMcpServer" /v ServerExePath /t REG_SZ /d "C:\Program Files\VBA MCP Server\VbaMcpServer.exe" /f
```

### Priority 3: Same Directory (Final Fallback)

**Path**: `{GUI_DIRECTORY}\VbaMcpServer.exe`

**When it works**:
- Both executables deployed together (e.g., portable installation)
- Manual xcopy deployment
- MSI installation (both files in `C:\Program Files\VBA MCP Server\`)
- Development builds using unified output directory (via Directory.Build.props)

**Example**:
```
C:\MyTools\
├── VbaMcpServer.exe        ← Server
├── VbaMcpServer.GUI.exe    ← GUI
└── appsettings.json
```

**Note for developers**: When building from source with the standard solution, both executables are output to the same `bin\{Configuration}\` directory due to Directory.Build.props settings, so this fallback works automatically for development builds.

## Implementation Details

### Code Flow

```csharp
private string FindMcpServerExecutable()
{
    List<string> candidates = [];

    // 1. Check appsettings.json
    var configPath = _configuration["VbaMcpServer:ServerExePath"];
    if (!string.IsNullOrWhiteSpace(configPath))
        candidates.Add(configPath);

    // 2. Check registry
    using var key = Registry.CurrentUser.OpenSubKey(@"Software\VbaMcpServer");
    var registryPath = key?.GetValue("ServerExePath") as string;
    if (!string.IsNullOrWhiteSpace(registryPath))
        candidates.Add(registryPath);

    // 3. Same directory (fallback)
    candidates.Add(Path.Combine(currentDir, "VbaMcpServer.exe"));

    // Search and return first match
    foreach (var candidate in candidates)
    {
        if (File.Exists(Path.GetFullPath(candidate)))
            return Path.GetFullPath(candidate);
    }

    return Path.Combine(currentDir, "VbaMcpServer.exe"); // Final fallback
}
```

### Logging

All path resolution attempts are logged to the Server Log tab in the GUI:

```
2025-12-26 10:15:23 - Using path from appsettings.json: C:\Custom\VbaMcpServer.exe
2025-12-26 10:15:23 - Current directory: C:\Program Files\VBA MCP Server\
2025-12-26 10:15:23 - Checking: C:\Custom\VbaMcpServer.exe
2025-12-26 10:15:23 - Found: C:\Custom\VbaMcpServer.exe
2025-12-26 10:15:23 - MCP Server path: C:\Custom\VbaMcpServer.exe
```

## Troubleshooting

### Problem: "MCP Server executable not found"

**Diagnosis steps**:

1. **Check the Server Log tab** for detailed search results
2. **Verify file exists**:
   ```powershell
   Test-Path "C:\Program Files\VBA MCP Server\VbaMcpServer.exe"
   ```
3. **Check registry**:
   ```powershell
   Get-ItemProperty -Path "HKCU:\Software\VbaMcpServer"
   ```
4. **Try manual configuration** in appsettings.json

### Problem: Wrong version being used

**Diagnosis**:
- Check log to see which path was selected
- Configuration/registry paths take priority over auto-detection

**Solution**:
- Remove or update appsettings.json
- Update registry value
- Rebuild the preferred version

### Problem: Development builds not found

**Possible causes**:
1. Building with custom output directories that override Directory.Build.props
2. Building projects independently instead of the whole solution
3. Output files in different directories

**Solutions**:
1. Build the entire solution (not individual projects) to use unified output:
   ```bash
   dotnet build -c Debug
   ```
2. Use appsettings.json to specify path manually:
   ```json
   {
     "VbaMcpServer": {
       "ServerExePath": "C:\\MyProjects\\vba-mcp\\src\\VbaMcpServer\\bin\\Debug\\VbaMcpServer.exe"
     }
   }
   ```
3. Ensure Directory.Build.props is present at solution root

## Best Practices

### For Developers

1. **No configuration needed**: Standard solution structure works automatically
2. **Switch configurations freely**: Debug/Release both auto-detected
3. **Override when needed**: Use appsettings.json for custom scenarios

### For End Users

1. **Use MSI installer**: Registry entry set automatically
2. **Portable deployment**: Copy both .exe files to same folder
3. **Custom location**: Edit appsettings.json (in same folder as GUI)

### For System Administrators

1. **Network deployment**: Use appsettings.json with UNC path:
   ```json
   {
     "VbaMcpServer": {
       "ServerExePath": "\\\\fileserver\\tools\\VbaMcpServer\\VbaMcpServer.exe"
     }
   }
   ```
2. **Group Policy deployment**: Set registry via GPO
3. **Citrix/RDS**: Ensure paths are accessible from user sessions

## Security Considerations

1. **No search PATH**: Executable is not searched in system PATH to avoid DLL hijacking
2. **Full path required**: Relative paths are resolved but absolute paths recommended
3. **User-specific registry**: Uses HKCU not HKLM to avoid privilege escalation
4. **Configuration file in application directory**: Prevents unauthorized modification

## Future Enhancements

Potential improvements for future versions:

1. **GUI Settings Dialog**: Visual path configuration instead of manual JSON editing
2. **Path Validation**: Check file signature/version before accepting
3. **Multiple Server Profiles**: Support different server configurations
4. **Auto-update Detection**: Detect and offer to use newer versions

---

<a name="japanese"></a>

## 概要

このドキュメントでは、VBA MCP Server Manager GUI が本番環境（インストール済み）と開発環境の両方で VbaMcpServer.exe 実行ファイルを見つける方法について説明します。

## 設計目標

1. **標準インストールでは設定不要**: MSI インストール後すぐに動作
2. **開発者フレンドリー**: 設定なしで開発ビルドを自動検出
3. **柔軟性**: カスタムシナリオ用のユーザーオーバーライドが可能
4. **堅牢性**: 優先方法が失敗した場合でも適切にフォールバック
5. **透過性**: デバッグ用にすべてのパス解決試行をログ記録

## 解決優先順位

GUI は以下の順序で VbaMcpServer.exe を検索します（最初に一致したものが使用されます）:

**注意**: 検索プロセスは3段階のみです。開発ビルドの自動検出は実装されていません。

### 優先順位 1: ユーザー設定（appsettings.json）

**ファイル**: `appsettings.json`（VbaMcpServer.GUI.exe と同じディレクトリ）

```json
{
  "VbaMcpServer": {
    "ServerExePath": "C:\\Custom\\Path\\VbaMcpServer.exe"
  }
}
```

**使用するタイミング**:
- カスタムインストールディレクトリ
- ネットワーク共有の場所
- 非標準的な配置
- 特殊な構造の開発環境

**使用例**:
```json
// カスタムインストール
{ "VbaMcpServer": { "ServerExePath": "D:\\Tools\\VbaMcpServer.exe" } }

// ネットワーク共有
{ "VbaMcpServer": { "ServerExePath": "\\\\server\\tools\\VbaMcpServer.exe" } }

// 相対パス（非推奨）
{ "VbaMcpServer": { "ServerExePath": "..\\..\\bin\\VbaMcpServer.exe" } }
```

### 優先順位 2: Windows レジストリ

**レジストリキー**: `HKEY_CURRENT_USER\Software\VbaMcpServer`
**値の名前**: `ServerExePath`
**値の種類**: `REG_SZ`
**値の例**: `C:\Program Files\VBA MCP Server\VbaMcpServer.exe`

**設定者**: MSI インストーラーがインストール時に設定
**使用するタイミング**: MSI による標準インストール

**手動でレジストリを設定する場合**:
```powershell
# PowerShell
New-Item -Path "HKCU:\Software\VbaMcpServer" -Force
Set-ItemProperty -Path "HKCU:\Software\VbaMcpServer" -Name "ServerExePath" -Value "C:\Program Files\VBA MCP Server\VbaMcpServer.exe"

# コマンドプロンプト
reg add "HKCU\Software\VbaMcpServer" /v ServerExePath /t REG_SZ /d "C:\Program Files\VBA MCP Server\VbaMcpServer.exe" /f
```

### 優先順位 3: 同じディレクトリ（最終フォールバック）

**パス**: `{GUI_DIRECTORY}\VbaMcpServer.exe`

**動作する場合**:
- 両方の実行ファイルが一緒に配置されている場合（ポータブルインストールなど）
- 手動での xcopy 配置
- MSI インストール（両方のファイルが `C:\Program Files\VBA MCP Server\` に配置）
- 統一出力ディレクトリを使用した開発ビルド（Directory.Build.props経由）

**例**:
```
C:\MyTools\
├── VbaMcpServer.exe        ← サーバー
├── VbaMcpServer.GUI.exe    ← GUI
└── appsettings.json
```

**開発者向け注意**: 標準ソリューションでソースからビルドする場合、Directory.Build.propsの設定により両方の実行ファイルが同じ`bin\{Configuration}\`ディレクトリに出力されるため、このフォールバックが開発ビルドでも自動的に機能します。

## 実装詳細

### コードフロー

```csharp
private string FindMcpServerExecutable()
{
    List<string> candidates = [];

    // 1. appsettings.json をチェック
    var configPath = _configuration["VbaMcpServer:ServerExePath"];
    if (!string.IsNullOrWhiteSpace(configPath))
        candidates.Add(configPath);

    // 2. レジストリをチェック
    using var key = Registry.CurrentUser.OpenSubKey(@"Software\VbaMcpServer");
    var registryPath = key?.GetValue("ServerExePath") as string;
    if (!string.IsNullOrWhiteSpace(registryPath))
        candidates.Add(registryPath);

    // 3. 同じディレクトリ（フォールバック）
    candidates.Add(Path.Combine(currentDir, "VbaMcpServer.exe"));

    // 検索して最初に一致したものを返す
    foreach (var candidate in candidates)
    {
        if (File.Exists(Path.GetFullPath(candidate)))
            return Path.GetFullPath(candidate);
    }

    return Path.Combine(currentDir, "VbaMcpServer.exe"); // 最終フォールバック
}
```

### ログ記録

すべてのパス解決試行は GUI の Server Log タブにログ記録されます:

```
2025-12-26 10:15:23 - Using path from appsettings.json: C:\Custom\VbaMcpServer.exe
2025-12-26 10:15:23 - Current directory: C:\Program Files\VBA MCP Server\
2025-12-26 10:15:23 - Checking: C:\Custom\VbaMcpServer.exe
2025-12-26 10:15:23 - Found: C:\Custom\VbaMcpServer.exe
2025-12-26 10:15:23 - MCP Server path: C:\Custom\VbaMcpServer.exe
```

## トラブルシューティング

### 問題: 「MCP Server executable not found」

**診断手順**:

1. **Server Log タブを確認** して詳細な検索結果を見る
2. **ファイルの存在を確認**:
   ```powershell
   Test-Path "C:\Program Files\VBA MCP Server\VbaMcpServer.exe"
   ```
3. **レジストリを確認**:
   ```powershell
   Get-ItemProperty -Path "HKCU:\Software\VbaMcpServer"
   ```
4. **appsettings.json で手動設定を試す**

### 問題: 間違ったバージョンが使用されている

**診断**:
- ログでどのパスが選択されたか確認
- 設定/レジストリのパスが自動検出より優先される

**解決策**:
- appsettings.json を削除または更新
- レジストリ値を更新
- 優先するバージョンをリビルド

### 問題: 開発ビルドが見つからない

**考えられる原因**:
1. Directory.Build.propsを上書きするカスタム出力ディレクトリでのビルド
2. ソリューション全体ではなくプロジェクトを個別にビルド
3. 出力ファイルが異なるディレクトリにある

**解決策**:
1. ソリューション全体をビルドして統一出力を使用:
   ```bash
   dotnet build -c Debug
   ```
2. appsettings.json を使用してパスを手動で指定:
   ```json
   {
     "VbaMcpServer": {
       "ServerExePath": "C:\\MyProjects\\vba-mcp\\src\\VbaMcpServer\\bin\\Debug\\VbaMcpServer.exe"
     }
   }
   ```
3. ソリューションルートにDirectory.Build.propsが存在することを確認

## ベストプラクティス

### 開発者向け

1. **設定不要**: 標準的なソリューション構造であれば自動で動作
2. **構成の自由な切り替え**: Debug/Release 両方が自動検出される
3. **必要に応じてオーバーライド**: カスタムシナリオには appsettings.json を使用

### エンドユーザー向け

1. **MSI インストーラーを使用**: レジストリエントリが自動設定される
2. **ポータブル配置**: 両方の .exe ファイルを同じフォルダにコピー
3. **カスタムの場所**: appsettings.json を編集（GUI と同じフォルダ）

### システム管理者向け

1. **ネットワーク配置**: appsettings.json で UNC パスを使用:
   ```json
   {
     "VbaMcpServer": {
       "ServerExePath": "\\\\fileserver\\tools\\VbaMcpServer\\VbaMcpServer.exe"
     }
   }
   ```
2. **グループポリシー配置**: GPO 経由でレジストリを設定
3. **Citrix/RDS**: パスがユーザーセッションからアクセス可能であることを確認

## セキュリティ上の考慮事項

1. **PATH を検索しない**: DLL ハイジャックを避けるため、システム PATH では実行ファイルを検索しない
2. **フルパスが必要**: 相対パスは解決されるが、絶対パスを推奨
3. **ユーザー固有のレジストリ**: 権限昇格を避けるため HKLM ではなく HKCU を使用
4. **アプリケーションディレクトリ内の設定ファイル**: 不正な変更を防止

## 将来の機能強化

将来のバージョンでの潜在的な改善点:

1. **GUI 設定ダイアログ**: 手動での JSON 編集ではなくビジュアルなパス設定
2. **パス検証**: 受け入れる前にファイル署名/バージョンをチェック
3. **複数のサーバープロファイル**: 異なるサーバー設定のサポート
4. **自動更新検出**: 新しいバージョンを検出して使用を提案
