# Configuration Guide / 設定ガイド

[English](#english) | [日本語](#japanese)

---

<a name="english"></a>

## GUI Manager Configuration

The VBA MCP Server Manager GUI uses multiple methods to locate the VbaMcpServer.exe file, checked in the following priority order:

### Priority 1: User Configuration File (Highest)

Edit `appsettings.json` located next to `VbaMcpServer.GUI.exe`:

```json
{
  "VbaMcpServer": {
    "ServerExePath": "C:\\Custom\\Path\\To\\VbaMcpServer.exe"
  }
}
```

**Use case**: Custom installation location, network path, or development environment with non-standard structure.

### Priority 2: Registry (Installer-set)

Automatically set during MSI installation:

```
HKEY_CURRENT_USER\Software\VbaMcpServer
  - ServerExePath: REG_SZ = "C:\Program Files\VBA MCP Server\VbaMcpServer.exe"
```

**Use case**: Standard installation via MSI installer.

### Priority 3: Same Directory

The GUI looks for `VbaMcpServer.exe` in the same directory as `VbaMcpServer.GUI.exe`.

**Use case**: Portable installation, deployed together.

### Priority 4: Development Build Detection (Legacy)

**Note**: As of version 0.2.0, all projects output to a unified `bin/` directory, so this priority is rarely used. It remains for backward compatibility.

Previously detected development builds in the following structure:

```
vba-mcp-server/
├── src/
│   ├── VbaMcpServer/bin/Debug/net8.0-windows/VbaMcpServer.exe
│   ├── VbaMcpServer/bin/Release/net8.0-windows/VbaMcpServer.exe
│   └── VbaMcpServer.GUI/bin/Debug/net8.0-windows/VbaMcpServer.GUI.exe
```

**Current behavior**: Projects now output to `bin/Debug/` or `bin/Release/` at the solution root, making Priority 3 (same directory) the standard for development.

## Troubleshooting

### GUI Cannot Find VbaMcpServer.exe

1. **Check Server Log tab** in the GUI for detailed path search information
2. **Verify file exists** at the expected location
3. **Override with appsettings.json** if automatic detection fails
4. **Check registry** (if installed via MSI):
   ```powershell
   Get-ItemProperty -Path "HKCU:\Software\VbaMcpServer"
   ```

### Custom Path Example

For development with custom directory structure:

**appsettings.json**:
```json
{
  "VbaMcpServer": {
    "ServerExePath": "C:\\MyProjects\\vba-tools\\VbaMcpServer.exe"
  }
}
```

### Network Path Example

For running the server from a network location:

**appsettings.json**:
```json
{
  "VbaMcpServer": {
    "ServerExePath": "\\\\fileserver\\tools\\VbaMcpServer\\VbaMcpServer.exe"
  }
}
```

## CLI Configuration

For direct MCP integration without GUI, configure Claude Desktop:

**%APPDATA%\Claude\claude_desktop_config.json**:
```json
{
  "mcpServers": {
    "vba": {
      "command": "C:\\Program Files\\VBA MCP Server\\VbaMcpServer.exe"
    }
  }
}
```

## Log Files

Logs are stored in user profile directory:

```
%USERPROFILE%\.vba-mcp-server\logs\
├── server\
│   └── server-2025-12-26.log
└── vba\
    └── vba-2025-12-26.log
```

Log rotation: Daily
Log format: JSON (Compact JSON format via Serilog)

---

<a name="japanese"></a>

## GUI マネージャー設定

VBA MCP Server Manager GUI は VbaMcpServer.exe ファイルを見つけるために複数の方法を使用し、以下の優先順位でチェックします:

### 優先順位 1: ユーザー設定ファイル（最優先）

`VbaMcpServer.GUI.exe` の隣にある `appsettings.json` を編集:

```json
{
  "VbaMcpServer": {
    "ServerExePath": "C:\\Custom\\Path\\To\\VbaMcpServer.exe"
  }
}
```

**使用例**: カスタムインストール場所、ネットワークパス、または非標準構造の開発環境。

### 優先順位 2: レジストリ（インストーラーが設定）

MSI インストール時に自動設定されます:

```
HKEY_CURRENT_USER\Software\VbaMcpServer
  - ServerExePath: REG_SZ = "C:\Program Files\VBA MCP Server\VbaMcpServer.exe"
```

**使用例**: MSI インストーラーによる標準インストール。

### 優先順位 3: 同じディレクトリ

GUI は `VbaMcpServer.GUI.exe` と同じディレクトリにある `VbaMcpServer.exe` を探します。

**使用例**: ポータブルインストール、一緒に配置された場合。

### 優先順位 4: 開発ビルド検出（レガシー）

**注意**: バージョン 0.2.0 以降、すべてのプロジェクトは統一された `bin/` ディレクトリに出力されるため、この優先順位はほとんど使用されません。後方互換性のために残されています。

以前は以下の構造で開発ビルドを検出していました:

```
vba-mcp-server/
├── src/
│   ├── VbaMcpServer/bin/Debug/net8.0-windows/VbaMcpServer.exe
│   ├── VbaMcpServer/bin/Release/net8.0-windows/VbaMcpServer.exe
│   └── VbaMcpServer.GUI/bin/Debug/net8.0-windows/VbaMcpServer.GUI.exe
```

**現在の動作**: プロジェクトはソリューションルートの `bin/Debug/` または `bin/Release/` に出力されるため、優先順位 3（同じディレクトリ）が開発時の標準となります。

## トラブルシューティング

### GUI が VbaMcpServer.exe を見つけられない

1. **Server Log タブを確認** して詳細なパス検索情報を確認
2. **ファイルの存在を確認** 期待される場所にファイルがあるか確認
3. **appsettings.json でオーバーライド** 自動検出が失敗する場合
4. **レジストリを確認**（MSI 経由でインストールした場合）:
   ```powershell
   Get-ItemProperty -Path "HKCU:\Software\VbaMcpServer"
   ```

### カスタムパスの例

カスタムディレクトリ構造での開発の場合:

**appsettings.json**:
```json
{
  "VbaMcpServer": {
    "ServerExePath": "C:\\MyProjects\\vba-tools\\VbaMcpServer.exe"
  }
}
```

### ネットワークパスの例

ネットワーク上の場所からサーバーを実行する場合:

**appsettings.json**:
```json
{
  "VbaMcpServer": {
    "ServerExePath": "\\\\fileserver\\tools\\VbaMcpServer\\VbaMcpServer.exe"
  }
}
```

## CLI 設定

GUI を使わずに直接 MCP 統合する場合は、Claude Desktop を設定します:

**%APPDATA%\Claude\claude_desktop_config.json**:
```json
{
  "mcpServers": {
    "vba": {
      "command": "C:\\Program Files\\VBA MCP Server\\VbaMcpServer.exe"
    }
  }
}
```

## ログファイル

ログはユーザープロファイルディレクトリに保存されます:

```
%USERPROFILE%\.vba-mcp-server\logs\
├── server\
│   └── server-2025-12-26.log
└── vba\
    └── vba-2025-12-26.log
```

ログローテーション: 日次
ログ形式: JSON（Serilog による Compact JSON 形式）
