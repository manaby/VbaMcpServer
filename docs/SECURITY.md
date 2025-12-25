# Office Security Settings / Officeセキュリティ設定

[English](#english) | [日本語](#japanese)

---

<a name="english"></a>

## Required Security Setting

To allow external applications to access VBA projects, you must enable a specific security setting in Microsoft Office.

### Why is this needed?

By default, Microsoft Office blocks programmatic access to VBA projects for security reasons. This protection prevents malicious software from modifying your macros. However, to use vba-mcp-server, you need to explicitly grant this access.

### How to Enable VBA Project Access

#### Excel

1. Open Excel
2. Click **File** → **Options**
3. Select **Trust Center** from the left panel
4. Click **Trust Center Settings...**
5. Select **Macro Settings** from the left panel
6. Check ☑ **Trust access to the VBA project object model**
7. Click **OK** to close Trust Center Settings
8. Click **OK** to close Excel Options

#### Access

1. Open Access
2. Click **File** → **Options**
3. Select **Trust Center** from the left panel
4. Click **Trust Center Settings...**
5. Select **Macro Settings** from the left panel
6. Check ☑ **Trust access to the VBA project object model**
7. Click **OK** to close Trust Center Settings
8. Click **OK** to close Access Options

### Group Policy (For IT Administrators)

If you need to deploy this setting across multiple machines, you can use Group Policy:

**Registry Key:**
```
HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Excel\Security
DWORD: AccessVBOM = 1
```

**PowerShell (per user):**
```powershell
# For Excel
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security" -Name "AccessVBOM" -Value 1 -Type DWord

# For Access
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Access\Security" -Name "AccessVBOM" -Value 1 -Type DWord
```

Note: Replace `16.0` with your Office version (15.0 for Office 2013, 16.0 for Office 2016/2019/365).

### Security Considerations

- Only enable this setting on machines where you trust the software that will access VBA projects
- This setting applies to ALL applications that attempt to access VBA projects, not just vba-mcp-server
- Consider disabling this setting when not actively using vba-mcp-server
- Always backup your files before allowing external modification of VBA code

---

<a name="japanese"></a>

## 必須のセキュリティ設定

外部アプリケーションが VBA プロジェクトにアクセスできるようにするには、Microsoft Office で特定のセキュリティ設定を有効にする必要があります。

### なぜ必要か？

Microsoft Office はセキュリティ上の理由から、デフォルトで VBA プロジェクトへのプログラムによるアクセスをブロックしています。この保護機能は、悪意のあるソフトウェアがマクロを改変するのを防ぎます。しかし、vba-mcp-server を使用するには、このアクセスを明示的に許可する必要があります。

### VBA プロジェクトアクセスを有効にする方法

#### Excel の場合

1. Excel を開く
2. **ファイル** → **オプション** をクリック
3. 左パネルから **トラストセンター** を選択
4. **トラストセンターの設定...** をクリック
5. 左パネルから **マクロの設定** を選択
6. ☑ **VBA プロジェクト オブジェクト モデルへのアクセスを信頼する** にチェック
7. **OK** をクリックしてトラストセンターの設定を閉じる
8. **OK** をクリックして Excel のオプションを閉じる

#### Access の場合

1. Access を開く
2. **ファイル** → **オプション** をクリック
3. 左パネルから **トラストセンター** を選択
4. **トラストセンターの設定...** をクリック
5. 左パネルから **マクロの設定** を選択
6. ☑ **VBA プロジェクト オブジェクト モデルへのアクセスを信頼する** にチェック
7. **OK** をクリックしてトラストセンターの設定を閉じる
8. **OK** をクリックして Access のオプションを閉じる

### グループポリシー（IT 管理者向け）

複数のマシンにこの設定を展開する必要がある場合は、グループポリシーを使用できます：

**レジストリキー:**
```
HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Excel\Security
DWORD: AccessVBOM = 1
```

**PowerShell（ユーザーごと）:**
```powershell
# Excel 用
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security" -Name "AccessVBOM" -Value 1 -Type DWord

# Access 用
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Access\Security" -Name "AccessVBOM" -Value 1 -Type DWord
```

注意: `16.0` はご使用の Office バージョンに置き換えてください（Office 2013 は 15.0、Office 2016/2019/365 は 16.0）。

### セキュリティに関する注意事項

- VBA プロジェクトにアクセスするソフトウェアを信頼できるマシンでのみ、この設定を有効にしてください
- この設定は vba-mcp-server だけでなく、VBA プロジェクトにアクセスしようとするすべてのアプリケーションに適用されます
- vba-mcp-server を積極的に使用していない場合は、この設定を無効にすることを検討してください
- VBA コードの外部変更を許可する前に、必ずファイルをバックアップしてください
