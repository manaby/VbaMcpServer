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

### Priority 3: Same Directory

**Path**: `{GUI_DIRECTORY}\VbaMcpServer.exe`

**When it works**:
- Both executables deployed together (e.g., portable installation)
- Manual xcopy deployment
- MSI installation (both files in `C:\Program Files\VBA MCP Server\`)

**Example**:
```
C:\MyTools\
├── VbaMcpServer.exe        ← Server
├── VbaMcpServer.GUI.exe    ← GUI
└── appsettings.json
```

### Priority 4: Development Build Detection

**Paths checked** (relative to GUI binary location):

```
GUI Location:
  src/VbaMcpServer.GUI/bin/{Debug|Release}/net8.0-windows/VbaMcpServer.GUI.exe

Searched Locations:
  ../../../../VbaMcpServer/bin/Debug/net8.0-windows/VbaMcpServer.exe
  ../../../../VbaMcpServer/bin/Release/net8.0-windows/VbaMcpServer.exe
  ../../../../VbaMcpServer/bin/Debug/net8.0-windows/win-x64/VbaMcpServer.exe
  ../../../../VbaMcpServer/bin/Release/net8.0-windows/win-x64/VbaMcpServer.exe
```

**When it works**:
- Standard Visual Studio solution structure
- Building from source
- Running/debugging in Visual Studio
- Both Debug and Release configurations supported
- Both normal builds and published builds supported

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

    // 3. Same directory
    candidates.Add(Path.Combine(currentDir, "VbaMcpServer.exe"));

    // 4. Development builds
    foreach (var config in new[] { "Debug", "Release" })
    {
        candidates.Add(Path.Combine(currentDir, "..", "..", "..", "..", "..", "VbaMcpServer", "bin", config, "net8.0-windows", "VbaMcpServer.exe"));
        candidates.Add(Path.Combine(currentDir, "..", "..", "..", "..", "..", "VbaMcpServer", "bin", config, "net8.0-windows", "win-x64", "VbaMcpServer.exe"));
    }

    // Search and return first match
    foreach (var candidate in candidates)
    {
        if (File.Exists(Path.GetFullPath(candidate)))
            return Path.GetFullPath(candidate);
    }

    return Path.Combine(currentDir, "VbaMcpServer.exe"); // Fallback
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

### Problem: Development builds not detected

**Possible causes**:
1. Non-standard solution structure
2. Building with custom output directories
3. Custom MSBuild properties

**Solutions**:
1. Use appsettings.json to specify path manually:
   ```json
   {
     "VbaMcpServer": {
       "ServerExePath": "C:\\MyProjects\\vba-mcp\\output\\VbaMcpServer.exe"
     }
   }
   ```
2. Adjust solution structure to match expected layout
3. Copy VbaMcpServer.exe to same directory as GUI

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

### 優先順位 3: 同じディレクトリ

**パス**: `{GUI_DIRECTORY}\VbaMcpServer.exe`

**動作する場合**:
- 両方の実行ファイルが一緒に配置されている場合（ポータブルインストールなど）
- 手動での xcopy 配置
- MSI インストール（両方のファイルが `C:\Program Files\VBA MCP Server\` に配置）

**例**:
```
C:\MyTools\
├── VbaMcpServer.exe        ← サーバー
├── VbaMcpServer.GUI.exe    ← GUI
└── appsettings.json
```

### 優先順位 4: 開発ビルド検出

**チェックされるパス**（GUI バイナリの場所からの相対パス）:

```
GUI の場所:
  src/VbaMcpServer.GUI/bin/{Debug|Release}/net8.0-windows/VbaMcpServer.GUI.exe

検索される場所:
  ../../../../VbaMcpServer/bin/Debug/net8.0-windows/VbaMcpServer.exe
  ../../../../VbaMcpServer/bin/Release/net8.0-windows/VbaMcpServer.exe
  ../../../../VbaMcpServer/bin/Debug/net8.0-windows/win-x64/VbaMcpServer.exe
  ../../../../VbaMcpServer/bin/Release/net8.0-windows/win-x64/VbaMcpServer.exe
```

**動作する場合**:
- 標準的な Visual Studio ソリューション構造
- ソースからのビルド
- Visual Studio での実行/デバッグ
- Debug と Release 両方の構成に対応
- 通常ビルドと発行ビルド両方に対応

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

    // 3. 同じディレクトリ
    candidates.Add(Path.Combine(currentDir, "VbaMcpServer.exe"));

    // 4. 開発ビルド
    foreach (var config in new[] { "Debug", "Release" })
    {
        candidates.Add(Path.Combine(currentDir, "..", "..", "..", "..", "..", "VbaMcpServer", "bin", config, "net8.0-windows", "VbaMcpServer.exe"));
        candidates.Add(Path.Combine(currentDir, "..", "..", "..", "..", "..", "VbaMcpServer", "bin", config, "net8.0-windows", "win-x64", "VbaMcpServer.exe"));
    }

    // 検索して最初に一致したものを返す
    foreach (var candidate in candidates)
    {
        if (File.Exists(Path.GetFullPath(candidate)))
            return Path.GetFullPath(candidate);
    }

    return Path.Combine(currentDir, "VbaMcpServer.exe"); // フォールバック
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

### 問題: 開発ビルドが検出されない

**考えられる原因**:
1. 非標準的なソリューション構造
2. カスタム出力ディレクトリでのビルド
3. カスタム MSBuild プロパティ

**解決策**:
1. appsettings.json を使用してパスを手動で指定:
   ```json
   {
     "VbaMcpServer": {
       "ServerExePath": "C:\\MyProjects\\vba-mcp\\output\\VbaMcpServer.exe"
     }
   }
   ```
2. ソリューション構造を期待されるレイアウトに調整
3. VbaMcpServer.exe を GUI と同じディレクトリにコピー

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
