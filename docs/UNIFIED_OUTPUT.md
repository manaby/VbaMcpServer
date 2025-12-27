# Unified Output Directory Structure / 統一出力ディレクトリ構成

[English](#english) | [日本語](#japanese)

---

<a name="english"></a>

## Overview

By unifying the output destination for all projects to the **`bin/` directory at the solution root**, we have achieved the same directory structure in both development and production environments.

## Directory Structure

### Before (Old Structure)

```
vba-mcp-server/
├── src/
│   ├── VbaMcpServer/
│   │   └── bin/Debug/net8.0-windows/
│   │       └── VbaMcpServer.exe          ← Scattered
│   ├── VbaMcpServer.GUI/
│   │   └── bin/Debug/net8.0-windows/
│   │       └── VbaMcpServer.GUI.exe      ← Scattered
│   └── VbaMcpServer.Core/
│       └── bin/Debug/net8.0-windows/
│           └── VbaMcpServer.Core.dll     ← Scattered
```

**Problems:**
- Complex relative path calculations needed for GUI to find server exe
- Different structures between development and production environments
- Path issues prone to occur during debugging

### After (New Unified Structure)

```
vba-mcp-server/
├── bin/
│   ├── Debug/                            ← Unified output directory
│   │   ├── VbaMcpServer.exe             ← All in same folder
│   │   ├── VbaMcpServer.GUI.exe         ← All in same folder
│   │   ├── VbaMcpServer.Core.dll        ← All in same folder
│   │   ├── appsettings.json             ← Config files too
│   │   └── (Other dependency DLLs)
│   └── Release/                          ← Release config same way
│       ├── VbaMcpServer.exe
│       ├── VbaMcpServer.GUI.exe
│       └── ...
└── src/
    ├── VbaMcpServer/
    ├── VbaMcpServer.GUI/
    ├── VbaMcpServer.Core/
    └── VbaMcpServer.Tests/
```

**Benefits:**
- ✅ Simple path resolution (just look in same directory)
- ✅ Same structure in development and production environments
- ✅ Easy debugging
- ✅ Distribution is just copying the folder

## Implementation

### Directory.Build.props

By placing `Directory.Build.props` at the solution root, we apply common settings to all projects.

**c:\\Users\\斎藤学\\OneDrive - 斉藤情報システムデザイン\\201_devs\\035_vba-mcp-server\\vba-mcp-server\\Directory.Build.props**:

```xml
<Project>
  <PropertyGroup>
    <!-- Unified output directory for all projects -->
    <BaseOutputPath>$(MSBuildThisFileDirectory)bin\</BaseOutputPath>
    <OutputPath>$(BaseOutputPath)$(Configuration)\</OutputPath>
    <AppendTargetFrameworkToOutputPath>false</AppendTargetFrameworkToOutputPath>
    <AppendRuntimeIdentifierToOutputPath>false</AppendRuntimeIdentifierToOutputPath>
  </PropertyGroup>
</Project>
```

**Property Meanings:**

| Property | Description | Effect |
|----------|-------------|--------|
| `BaseOutputPath` | Base output path | `{SolutionRoot}/bin/` |
| `OutputPath` | Final output path | `{SolutionRoot}/bin/{Debug\|Release}/` |
| `AppendTargetFrameworkToOutputPath` | Don't append framework name | `false` → Don't add `net8.0-windows/` |
| `AppendRuntimeIdentifierToOutputPath` | Don't append Runtime ID | `false` → Don't add `win-x64/` |

### Simplified Path Resolution

The `FindMcpServerExecutable()` method in **MainForm.cs** is now dramatically simpler:

**Before (Complex relative path calculation)**:
```csharp
// Go up 5 levels and search another project's bin folder...
candidates.Add(Path.Combine(currentDir, "..", "..", "..", "..", "..",
    "VbaMcpServer", "bin", config, "net8.0-windows", "VbaMcpServer.exe"));
```

**After (Just look in same directory)**:
```csharp
// Just look in same directory
var sameDirPath = Path.Combine(currentDir, "VbaMcpServer.exe");
```

## Build Behavior

### Debug Build

```bash
dotnet build
# or
dotnet build -c Debug
```

**Output to**: `bin/Debug/`

### Release Build

```bash
dotnet build -c Release
```

**Output to**: `bin/Release/`

### Clean Build

```bash
dotnet clean
dotnet build
```

The old `src/{project}/bin/` is deleted, and only the new `bin/` is used.

## Distribution

### Distributing Development Builds

```bash
# Copy entire bin/Debug/ or bin/Release/ folder
xcopy bin\Release\*.* D:\Distribution\ /S /E
```

### Publish Build (Single exe file)

```bash
dotnet publish src/VbaMcpServer -c Release -r win-x64 --self-contained /p:PublishSingleFile=true
dotnet publish src/VbaMcpServer.GUI -c Release -r win-x64 --self-contained /p:PublishSingleFile=true
```

**Output to**: Each project's `bin/Release/win-x64/publish/`

**Note**: Publish builds are not affected by the unified output directory (they use project-specific output paths).

## Impact on Installer

WiX installer configuration is also simplified:

**Before (Complex relative paths)**:
```xml
<File Source="../src/VbaMcpServer/bin/Release/net8.0-windows/win-x64/publish/VbaMcpServer.exe" />
<File Source="../src/VbaMcpServer.GUI/bin/Release/net8.0-windows/win-x64/publish/VbaMcpServer.GUI.exe" />
```

**After (Publish version as before)**:
```xml
<!-- Publish version uses project-specific paths as before -->
<File Source="../src/VbaMcpServer/bin/Release/win-x64/publish/VbaMcpServer.exe" />
<File Source="../src/VbaMcpServer.GUI/bin/Release/win-x64/publish/VbaMcpServer.GUI.exe" />
```

Or, if using regular builds:
```xml
<!-- Regular build version from unified directory -->
<File Source="../bin/Release/VbaMcpServer.exe" />
<File Source="../bin/Release/VbaMcpServer.GUI.exe" />
```

## Troubleshooting

### Old Build Files Remain

**Symptom**: Files remain in old `src/{project}/bin/`

**Solution**:
```bash
dotnet clean
# or
git clean -fdx bin/
```

### Visual Studio Using Cache

**Symptom**: After building, files don't appear in `bin/Debug/`

**Solution**:
1. Close Visual Studio
2. Delete `bin/` folder
3. Delete `.vs/` folder (hidden folder)
4. Reopen Visual Studio and rebuild

### Tests Not Found

**Symptom**: Tests don't run with `dotnet test`

**Solution**:
Test assemblies are also placed in the unified output directory, so they should run normally:
```bash
dotnet test
```

## Summary

The unified output directory provides:

- ✅ Dramatically simplified path resolution logic
- ✅ Easier debugging in development environment
- ✅ Development and production environments behave identically
- ✅ Distribution is just copying a folder
- ✅ Future extensions are easier

This change fundamentally solved the "server not found in development environment" issue pointed out by users.

---

<a name="japanese"></a>

## 概要

すべてのプロジェクトの出力先を **ソリューションルートの `bin/` ディレクトリ** に統一することで、開発環境と本番環境で同じディレクトリ構造を実現しました。

## ディレクトリ構造

### Before（以前の構造）

```
vba-mcp-server/
├── src/
│   ├── VbaMcpServer/
│   │   └── bin/Debug/net8.0-windows/
│   │       └── VbaMcpServer.exe          ← バラバラ
│   ├── VbaMcpServer.GUI/
│   │   └── bin/Debug/net8.0-windows/
│   │       └── VbaMcpServer.GUI.exe      ← バラバラ
│   └── VbaMcpServer.Core/
│       └── bin/Debug/net8.0-windows/
│           └── VbaMcpServer.Core.dll     ← バラバラ
```

**問題点:**
- GUIがサーバーのexeを見つけるために複雑な相対パス計算が必要
- 開発環境と本番環境で構造が異なる
- デバッグ時にパスの問題が発生しやすい

### After（新しい統一構造）

```
vba-mcp-server/
├── bin/
│   ├── Debug/                            ← 統一された出力ディレクトリ
│   │   ├── VbaMcpServer.exe             ← すべて同じフォルダ
│   │   ├── VbaMcpServer.GUI.exe         ← すべて同じフォルダ
│   │   ├── VbaMcpServer.Core.dll        ← すべて同じフォルダ
│   │   ├── appsettings.json             ← 設定ファイルも同じフォルダ
│   │   └── (その他の依存DLL)
│   └── Release/                          ← Release構成も同様
│       ├── VbaMcpServer.exe
│       ├── VbaMcpServer.GUI.exe
│       └── ...
└── src/
    ├── VbaMcpServer/
    ├── VbaMcpServer.GUI/
    ├── VbaMcpServer.Core/
    └── VbaMcpServer.Tests/
```

**利点:**
- ✅ パス解決が単純（同じディレクトリを見るだけ）
- ✅ 開発環境と本番環境で同じ構造
- ✅ デバッグが容易
- ✅ 配布時にフォルダごとコピーするだけ

## 実装方法

### Directory.Build.props

ソリューションルートに `Directory.Build.props` を配置することで、すべてのプロジェクトに共通の設定を適用します。

**c:\\Users\\斎藤学\\OneDrive - 斉藤情報システムデザイン\\201_devs\\035_vba-mcp-server\\vba-mcp-server\\Directory.Build.props**:

```xml
<Project>
  <PropertyGroup>
    <!-- Unified output directory for all projects -->
    <BaseOutputPath>$(MSBuildThisFileDirectory)bin\</BaseOutputPath>
    <OutputPath>$(BaseOutputPath)$(Configuration)\</OutputPath>
    <AppendTargetFrameworkToOutputPath>false</AppendTargetFrameworkToOutputPath>
    <AppendRuntimeIdentifierToOutputPath>false</AppendRuntimeIdentifierToOutputPath>
  </PropertyGroup>
</Project>
```

**設定の意味:**

| プロパティ | 説明 | 効果 |
|-----------|------|------|
| `BaseOutputPath` | ベース出力パス | `{ソリューションルート}/bin/` |
| `OutputPath` | 最終出力パス | `{ソリューションルート}/bin/{Debug\|Release}/` |
| `AppendTargetFrameworkToOutputPath` | フレームワーク名を追加しない | `false` → `net8.0-windows/` を追加しない |
| `AppendRuntimeIdentifierToOutputPath` | Runtime IDを追加しない | `false` → `win-x64/` を追加しない |

### パス解決の簡素化

**MainForm.cs** の `FindMcpServerExecutable()` メソッドが大幅にシンプルになりました:

**Before（複雑な相対パス計算）**:
```csharp
// 5階層上がって、別のプロジェクトのbinフォルダを探す...
candidates.Add(Path.Combine(currentDir, "..", "..", "..", "..", "..",
    "VbaMcpServer", "bin", config, "net8.0-windows", "VbaMcpServer.exe"));
```

**After（同じディレクトリを見るだけ）**:
```csharp
// 同じディレクトリを見るだけ
var sameDirPath = Path.Combine(currentDir, "VbaMcpServer.exe");
```

## ビルド動作

### Debug ビルド

```bash
dotnet build
# または
dotnet build -c Debug
```

**出力先**: `bin/Debug/`

### Release ビルド

```bash
dotnet build -c Release
```

**出力先**: `bin/Release/`

### クリーンビルド

```bash
dotnet clean
dotnet build
```

旧来の `src/{プロジェクト}/bin/` は削除され、新しい `bin/` のみが使用されます。

## 配布方法

### 開発ビルドの配布

```bash
# bin/Debug/ または bin/Release/ フォルダごとコピー
xcopy bin\Release\*.* D:\Distribution\ /S /E
```

### Publish ビルド（単一exeファイル）

```bash
dotnet publish src/VbaMcpServer -c Release -r win-x64 --self-contained /p:PublishSingleFile=true
dotnet publish src/VbaMcpServer.GUI -c Release -r win-x64 --self-contained /p:PublishSingleFile=true
```

**出力先**: 各プロジェクトの `bin/Release/win-x64/publish/`

**注意**: Publishビルドは統一出力ディレクトリの影響を受けません（プロジェクト個別の出力パスを使用）。

## インストーラーへの影響

WiX インストーラーの構成も簡素化されます:

**Before（複雑な相対パス）**:
```xml
<File Source="../src/VbaMcpServer/bin/Release/net8.0-windows/win-x64/publish/VbaMcpServer.exe" />
<File Source="../src/VbaMcpServer.GUI/bin/Release/net8.0-windows/win-x64/publish/VbaMcpServer.GUI.exe" />
```

**After（Publishビルドは従来通り）**:
```xml
<!-- Publish版は従来通りプロジェクト個別のパス -->
<File Source="../src/VbaMcpServer/bin/Release/win-x64/publish/VbaMcpServer.exe" />
<File Source="../src/VbaMcpServer.GUI/bin/Release/win-x64/publish/VbaMcpServer.GUI.exe" />
```

または、通常ビルドを使用する場合:
```xml
<!-- 通常ビルド版は統一ディレクトリから -->
<File Source="../bin/Release/VbaMcpServer.exe" />
<File Source="../bin/Release/VbaMcpServer.GUI.exe" />
```

## トラブルシューティング

### 古いビルドファイルが残っている

**症状**: 古い `src/{プロジェクト}/bin/` にファイルが残っている

**解決策**:
```bash
dotnet clean
# または
git clean -fdx bin/
```

### Visual Studio がキャッシュを使用している

**症状**: ビルドしても `bin/Debug/` に出力されない

**解決策**:
1. Visual Studio を閉じる
2. `bin/` フォルダを削除
3. `.vs/` フォルダを削除（隠しフォルダ）
4. Visual Studio を開き直してリビルド

### テストが見つからない

**症状**: `dotnet test` でテストが実行されない

**解決策**:
統一出力ディレクトリにテストアセンブリも配置されているため、通常通り実行可能:
```bash
dotnet test
```

## まとめ

統一出力ディレクトリにより:

- ✅ パス解決ロジックが劇的にシンプルになった
- ✅ 開発環境でのデバッグが容易になった
- ✅ 本番環境と開発環境の動作が一致
- ✅ 配布がフォルダコピーだけで完結
- ✅ 将来的な拡張が容易

この変更により、ユーザーが指摘した「開発環境でサーバーが見つからない」問題が根本的に解決されました。
