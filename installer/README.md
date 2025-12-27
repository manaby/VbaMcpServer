# VBA MCP Server Installer

このディレクトリには、VBA MCP ServerのWindowsインストーラ(MSI)を作成するためのファイルが含まれています。

## 前提条件

- WiX Toolset v5.0+ がインストールされていること
  ```bash
  dotnet tool install --global wix
  ```

## インストーラのビルド

### Windowsの場合

```cmd
cd installer
build-installer.bat
```

### クロスプラットフォーム(PowerShell)の場合

```powershell
# 1. アプリケーションをPublish
dotnet publish ../src/VbaMcpServer -c Release -r win-x64 --self-contained /p:PublishSingleFile=true
dotnet publish ../src/VbaMcpServer.GUI -c Release -r win-x64 --self-contained /p:PublishSingleFile=true

# 2. MSIをビルド
cd installer
wix build Product.wxs -o bin/VbaMcpServer.msi -arch x64
```

## 出力

ビルドが成功すると、以下のファイルが生成されます:

- `installer/bin/VbaMcpServer.msi` - インストーラ本体

## インストール

1. `VbaMcpServer.msi` をダブルクリック
2. インストールウィザードに従って進む
3. スタートメニューから「VBA MCP Server Manager」を起動

## アンインストール

- Windowsの「設定」→「アプリ」→「VBA MCP Server」を選択してアンインストール
- または、コントロールパネルの「プログラムと機能」から削除

## ファイル構成

- `Product.wxs` - WiXインストーラ定義ファイル
- `build-installer.bat` - Windowsビルドスクリプト
- `README.md` - このファイル

## インストール先

- プログラムファイル: `C:\Program Files\VBA MCP Server\`
- スタートメニュー: `VBA MCP Server\VBA MCP Server Manager`
- レジストリ: `HKCU\Software\VbaMcpServer`
