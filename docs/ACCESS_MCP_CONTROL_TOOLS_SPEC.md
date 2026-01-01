# Access MCP サーバー拡張仕様書
# フォーム・レポート コントロール操作機能

## 1. 概要

本仕様書は、Access MCP サーバーに追加するフォームおよびレポートのコントロール操作機能について定義する。

### 1.1 目的

- Accessフォーム/レポート上のコントロール一覧を取得可能にする
- コントロールのプロパティを読み取り・設定可能にする
- AI支援によるフォーム/レポート開発を効率化する

### 1.2 追加ツール一覧

| ツール名 | 種別 | 説明 |
|----------|------|------|
| `get_access_form_controls` | 読取 | フォームのコントロール一覧を取得 |
| `get_access_form_control_properties` | 読取 | フォームコントロールのプロパティを取得 |
| `set_access_form_control_property` | 書込 | フォームコントロールのプロパティを設定 |
| `get_access_report_controls` | 読取 | レポートのコントロール一覧を取得 |
| `get_access_report_control_properties` | 読取 | レポートコントロールのプロパティを取得 |
| `set_access_report_control_property` | 書込 | レポートコントロールのプロパティを設定 |

---

## 2. フォーム関連ツール

### 2.1 get_access_form_controls

フォーム上のすべてのコントロールの一覧を取得する。

#### パラメータ

| パラメータ | 型 | 必須 | デフォルト | 説明 |
|------------|------|------|------------|------|
| `filePath` | string | ✓ | - | Accessデータベースのフルパス |
| `formName` | string | ✓ | - | フォーム名 |
| `includeChildren` | boolean | | false | サブフォーム内のコントロールも含めるか |
| `format` | string | | "json" | 出力形式: "json" または "csv" |

#### レスポンス例

```json
{
  "file": "C:\\Users\\user\\Desktop\\sample.accdb",
  "formName": "フォーム1",
  "controlCount": 5,
  "controls": [
    {
      "name": "ラベル0",
      "controlType": "Label",
      "controlTypeId": 100,
      "section": "Detail",
      "sectionId": 0,
      "left": 300,
      "top": 200,
      "width": 1500,
      "height": 300,
      "visible": true,
      "enabled": true,
      "tabIndex": null,
      "controlSource": null,
      "parent": null
    },
    {
      "name": "テキスト1",
      "controlType": "TextBox",
      "controlTypeId": 109,
      "section": "Detail",
      "sectionId": 0,
      "left": 1900,
      "top": 200,
      "width": 2400,
      "height": 300,
      "visible": true,
      "enabled": true,
      "tabIndex": 0,
      "controlSource": "フィールド1",
      "parent": null
    },
    {
      "name": "コマンド2",
      "controlType": "CommandButton",
      "controlTypeId": 104,
      "section": "Detail",
      "sectionId": 0,
      "left": 300,
      "top": 800,
      "width": 1500,
      "height": 450,
      "visible": true,
      "enabled": true,
      "tabIndex": 1,
      "controlSource": null,
      "parent": null
    },
    {
      "name": "サブフォーム3",
      "controlType": "SubForm",
      "controlTypeId": 112,
      "section": "Detail",
      "sectionId": 0,
      "left": 300,
      "top": 1500,
      "width": 5000,
      "height": 2000,
      "visible": true,
      "enabled": true,
      "tabIndex": 2,
      "controlSource": null,
      "parent": null,
      "sourceObject": "サブフォーム1"
    }
  ]
}
```

#### コントロールタイプID（主要なもの）

| ID | タイプ名 | 説明 |
|----|----------|------|
| 100 | Label | ラベル |
| 101 | Rectangle | 四角形 |
| 102 | Line | 線 |
| 103 | Image | イメージ |
| 104 | CommandButton | コマンドボタン |
| 105 | OptionButton | オプションボタン |
| 106 | CheckBox | チェックボックス |
| 107 | OptionGroup | オプショングループ |
| 108 | BoundObjectFrame | 連結オブジェクトフレーム |
| 109 | TextBox | テキストボックス |
| 110 | ListBox | リストボックス |
| 111 | ComboBox | コンボボックス |
| 112 | SubForm | サブフォーム/サブレポート |
| 114 | ObjectFrame | 非連結オブジェクトフレーム |
| 118 | PageBreak | 改ページ |
| 119 | CustomControl | ActiveXコントロール |
| 122 | ToggleButton | トグルボタン |
| 123 | TabControl | タブコントロール |
| 124 | Page | ページ（タブ内） |
| 126 | Attachment | 添付ファイル |
| 127 | NavigationControl | ナビゲーションコントロール |
| 128 | NavigationButton | ナビゲーションボタン |
| 129 | WebBrowserControl | Webブラウザコントロール |

#### セクションID

| ID | セクション名 |
|----|--------------|
| 0 | Detail（詳細） |
| 1 | Header（フォームヘッダー） |
| 2 | Footer（フォームフッター） |
| 3 | PageHeader（ページヘッダー） |
| 4 | PageFooter（ページフッター） |

---

### 2.2 get_access_form_control_properties

指定したコントロールの詳細プロパティを取得する。

#### パラメータ

| パラメータ | 型 | 必須 | デフォルト | 説明 |
|------------|------|------|------------|------|
| `filePath` | string | ✓ | - | Accessデータベースのフルパス |
| `formName` | string | ✓ | - | フォーム名 |
| `controlName` | string | ✓ | - | コントロール名 |
| `properties` | string[] | | null | 取得するプロパティ名の配列（null=全て） |

#### レスポンス例（テキストボックスの場合）

```json
{
  "file": "C:\\Users\\user\\Desktop\\sample.accdb",
  "formName": "フォーム1",
  "controlName": "テキスト1",
  "controlType": "TextBox",
  "properties": {
    "Name": "テキスト1",
    "ControlSource": "フィールド1",
    "Format": "",
    "DecimalPlaces": 255,
    "InputMask": "",
    "DefaultValue": "",
    "ValidationRule": "",
    "ValidationText": "",
    "StatusBarText": "",
    "Visible": true,
    "Enabled": true,
    "Locked": false,
    "TabStop": true,
    "TabIndex": 0,
    "Left": 1900,
    "Top": 200,
    "Width": 2400,
    "Height": 300,
    "BackStyle": 1,
    "BackColor": 16777215,
    "BorderStyle": 1,
    "BorderColor": 0,
    "BorderWidth": 0,
    "ForeColor": 0,
    "FontName": "Yu Gothic UI",
    "FontSize": 11,
    "FontWeight": 400,
    "FontItalic": false,
    "FontUnderline": false,
    "TextAlign": 0,
    "ScrollBars": 0,
    "CanGrow": false,
    "CanShrink": false,
    "EnterKeyBehavior": false,
    "AllowAutoCorrect": true,
    "Tag": "",
    "ControlTipText": "",
    "ShortcutMenuBar": ""
  }
}
```

#### レスポンス例（コマンドボタンの場合）

```json
{
  "file": "C:\\Users\\user\\Desktop\\sample.accdb",
  "formName": "フォーム1",
  "controlName": "コマンド2",
  "controlType": "CommandButton",
  "properties": {
    "Name": "コマンド2",
    "Caption": "実行",
    "Picture": "",
    "PictureType": 0,
    "Visible": true,
    "Enabled": true,
    "TabStop": true,
    "TabIndex": 1,
    "Left": 300,
    "Top": 800,
    "Width": 1500,
    "Height": 450,
    "BackColor": 15064278,
    "BackStyle": 1,
    "BorderStyle": 1,
    "ForeColor": 0,
    "FontName": "Yu Gothic UI",
    "FontSize": 11,
    "FontWeight": 400,
    "Transparent": false,
    "HoverColor": 16772300,
    "HoverForeColor": 0,
    "PressedColor": 13395711,
    "PressedForeColor": 0,
    "Tag": "",
    "ControlTipText": "",
    "HyperlinkAddress": "",
    "OnClick": "[Event Procedure]",
    "OnEnter": "",
    "OnExit": "",
    "OnGotFocus": "",
    "OnLostFocus": ""
  }
}
```

#### 特定プロパティのみ取得する例

リクエスト:
```json
{
  "filePath": "C:\\Users\\user\\Desktop\\sample.accdb",
  "formName": "フォーム1",
  "controlName": "テキスト1",
  "properties": ["ControlSource", "Visible", "Enabled", "Left", "Top", "Width", "Height"]
}
```

レスポンス:
```json
{
  "file": "C:\\Users\\user\\Desktop\\sample.accdb",
  "formName": "フォーム1",
  "controlName": "テキスト1",
  "controlType": "TextBox",
  "properties": {
    "ControlSource": "フィールド1",
    "Visible": true,
    "Enabled": true,
    "Left": 1900,
    "Top": 200,
    "Width": 2400,
    "Height": 300
  }
}
```

---

### 2.3 set_access_form_control_property

指定したコントロールのプロパティを設定する。

#### パラメータ

| パラメータ | 型 | 必須 | デフォルト | 説明 |
|------------|------|------|------------|------|
| `filePath` | string | ✓ | - | Accessデータベースのフルパス |
| `formName` | string | ✓ | - | フォーム名 |
| `controlName` | string | ✓ | - | コントロール名 |
| `propertyName` | string | ✓ | - | 設定するプロパティ名 |
| `propertyValue` | any | ✓ | - | 設定する値 |

#### リクエスト例

```json
{
  "filePath": "C:\\Users\\user\\Desktop\\sample.accdb",
  "formName": "フォーム1",
  "controlName": "テキスト1",
  "propertyName": "BackColor",
  "propertyValue": 16777164
}
```

#### レスポンス例

```json
{
  "success": true,
  "file": "C:\\Users\\user\\Desktop\\sample.accdb",
  "formName": "フォーム1",
  "controlName": "テキスト1",
  "propertyName": "BackColor",
  "previousValue": 16777215,
  "newValue": 16777164
}
```

#### エラーレスポンス例

```json
{
  "success": false,
  "error": "Property 'InvalidProperty' not found on control 'テキスト1'",
  "errorCode": "PROPERTY_NOT_FOUND"
}
```

#### 注意事項

- フォームはデザインビューで開かれる
- 読み取り専用プロパティは設定不可（エラーを返す）
- 設定後、フォームは保存される
- 一部のプロパティはフォームを閉じた状態でのみ変更可能

---

## 3. レポート関連ツール

### 3.1 get_access_report_controls

レポート上のすべてのコントロールの一覧を取得する。

#### パラメータ

| パラメータ | 型 | 必須 | デフォルト | 説明 |
|------------|------|------|------------|------|
| `filePath` | string | ✓ | - | Accessデータベースのフルパス |
| `reportName` | string | ✓ | - | レポート名 |
| `includeChildren` | boolean | | false | サブレポート内のコントロールも含めるか |
| `format` | string | | "json" | 出力形式: "json" または "csv" |

#### レスポンス例

```json
{
  "file": "C:\\Users\\user\\Desktop\\sample.accdb",
  "reportName": "レポート1",
  "controlCount": 4,
  "controls": [
    {
      "name": "レポートヘッダーラベル",
      "controlType": "Label",
      "controlTypeId": 100,
      "section": "ReportHeader",
      "sectionId": 3,
      "left": 0,
      "top": 0,
      "width": 5000,
      "height": 600,
      "visible": true,
      "controlSource": null,
      "parent": null
    },
    {
      "name": "フィールド1",
      "controlType": "TextBox",
      "controlTypeId": 109,
      "section": "Detail",
      "sectionId": 0,
      "left": 0,
      "top": 0,
      "width": 2400,
      "height": 300,
      "visible": true,
      "controlSource": "フィールド1",
      "parent": null
    }
  ]
}
```

#### レポートのセクションID

| ID | セクション名 |
|----|--------------|
| 0 | Detail（詳細） |
| 1 | Header（レポートヘッダー）※実際は3 |
| 2 | Footer（レポートフッター）※実際は4 |
| 3 | PageHeader（ページヘッダー）※実際は1 |
| 4 | PageFooter（ページフッター）※実際は2 |
| 5 | GroupHeader（グループヘッダー） |
| 6 | GroupFooter（グループフッター） |

※ Accessの内部実装により、セクションIDの割り当ては非直感的な場合がある。
  実装時は `Section` プロパティの実際の値を確認すること。

---

### 3.2 get_access_report_control_properties

指定したレポートコントロールの詳細プロパティを取得する。

#### パラメータ

| パラメータ | 型 | 必須 | デフォルト | 説明 |
|------------|------|------|------------|------|
| `filePath` | string | ✓ | - | Accessデータベースのフルパス |
| `reportName` | string | ✓ | - | レポート名 |
| `controlName` | string | ✓ | - | コントロール名 |
| `properties` | string[] | | null | 取得するプロパティ名の配列（null=全て） |

#### レスポンス例

```json
{
  "file": "C:\\Users\\user\\Desktop\\sample.accdb",
  "reportName": "レポート1",
  "controlName": "フィールド1",
  "controlType": "TextBox",
  "properties": {
    "Name": "フィールド1",
    "ControlSource": "フィールド1",
    "Format": "",
    "DecimalPlaces": 255,
    "Visible": true,
    "Left": 0,
    "Top": 0,
    "Width": 2400,
    "Height": 300,
    "BackStyle": 0,
    "BackColor": 16777215,
    "BorderStyle": 0,
    "ForeColor": 0,
    "FontName": "Yu Gothic UI",
    "FontSize": 11,
    "FontWeight": 400,
    "CanGrow": true,
    "CanShrink": false,
    "RunningSum": 0,
    "HideDuplicates": false,
    "Tag": ""
  }
}
```

---

### 3.3 set_access_report_control_property

指定したレポートコントロールのプロパティを設定する。

#### パラメータ

| パラメータ | 型 | 必須 | デフォルト | 説明 |
|------------|------|------|------------|------|
| `filePath` | string | ✓ | - | Accessデータベースのフルパス |
| `reportName` | string | ✓ | - | レポート名 |
| `controlName` | string | ✓ | - | コントロール名 |
| `propertyName` | string | ✓ | - | 設定するプロパティ名 |
| `propertyValue` | any | ✓ | - | 設定する値 |

#### リクエスト例

```json
{
  "filePath": "C:\\Users\\user\\Desktop\\sample.accdb",
  "reportName": "レポート1",
  "controlName": "フィールド1",
  "propertyName": "CanGrow",
  "propertyValue": true
}
```

#### レスポンス例

```json
{
  "success": true,
  "file": "C:\\Users\\user\\Desktop\\sample.accdb",
  "reportName": "レポート1",
  "controlName": "フィールド1",
  "propertyName": "CanGrow",
  "previousValue": false,
  "newValue": true
}
```

---

## 4. 実装上の考慮事項

### 4.1 フォーム/レポートのオープンモード

コントロールのプロパティにアクセスするには、フォーム/レポートをデザインビューで開く必要がある。

```csharp
// C# 実装例
accessApp.DoCmd.OpenForm(formName, AcFormView.acDesign);
// または
accessApp.DoCmd.OpenReport(reportName, AcView.acViewDesign);
```

### 4.2 変更の保存

プロパティを変更した後は、明示的に保存する必要がある。

```csharp
// C# 実装例
accessApp.DoCmd.Save(AcObjectType.acForm, formName);
accessApp.DoCmd.Close(AcObjectType.acForm, formName, AcCloseSave.acSaveYes);
```

### 4.3 エラーハンドリング

| エラーコード | 説明 |
|--------------|------|
| `FORM_NOT_FOUND` | 指定されたフォームが存在しない |
| `REPORT_NOT_FOUND` | 指定されたレポートが存在しない |
| `CONTROL_NOT_FOUND` | 指定されたコントロールが存在しない |
| `PROPERTY_NOT_FOUND` | 指定されたプロパティが存在しない |
| `PROPERTY_READ_ONLY` | プロパティが読み取り専用 |
| `INVALID_VALUE` | プロパティに無効な値が指定された |
| `FORM_IN_USE` | フォームが既に使用中（排他エラー） |
| `ACCESS_DENIED` | アクセス権限がない |

### 4.4 COM型の変換

Accessのプロパティ値は様々な型で返されるため、適切な変換が必要。

| Access型 | C#型 | JSON型 |
|----------|------|--------|
| Integer | int | number |
| Long | int | number |
| Single | float | number |
| Double | double | number |
| String | string | string |
| Boolean | bool | boolean |
| Date | DateTime | string (ISO 8601) |
| Currency | decimal | number |
| Null | null | null |
| OLE Color | int | number |

### 4.5 パフォーマンス考慮

- コントロール一覧取得時は、必要最小限のプロパティのみ取得する
- 大量のプロパティ取得時はバッチ処理を検討
- フォーム/レポートの開閉は最小限に抑える

---

## 5. 将来の拡張案

### 5.1 コントロールの作成・削除

```
create_access_form_control     - フォームにコントロールを追加
delete_access_form_control     - フォームからコントロールを削除
create_access_report_control   - レポートにコントロールを追加
delete_access_report_control   - レポートからコントロールを削除
```

### 5.2 一括プロパティ設定

```
set_access_form_control_properties   - 複数プロパティを一括設定
set_access_report_control_properties - 複数プロパティを一括設定
```

### 5.3 フォーム/レポートのセクション操作

```
get_access_form_sections       - フォームセクション情報取得
set_access_form_section_property   - セクションプロパティ設定
get_access_report_sections     - レポートセクション情報取得
set_access_report_section_property - セクションプロパティ設定
```

### 5.4 フォーム/レポート自体のプロパティ操作

```
get_access_form_properties     - フォーム自体のプロパティ取得
set_access_form_property       - フォーム自体のプロパティ設定
get_access_report_properties   - レポート自体のプロパティ取得
set_access_report_property     - レポート自体のプロパティ設定
```

---

## 6. 実装優先度

| 優先度 | ツール | 理由 |
|--------|--------|------|
| 高 | `get_access_form_controls` | 基本機能として必須 |
| 高 | `get_access_form_control_properties` | コントロール詳細取得に必須 |
| 高 | `get_access_report_controls` | レポート対応に必須 |
| 高 | `get_access_report_control_properties` | レポートコントロール詳細取得に必須 |
| 中 | `set_access_form_control_property` | プロパティ変更機能 |
| 中 | `set_access_report_control_property` | プロパティ変更機能 |
| 低 | コントロール作成・削除 | 高度な機能 |
| 低 | 一括プロパティ設定 | 効率化機能 |

---

## 7. テスト計画

### 7.1 基本テスト

- [ ] 空のフォーム/レポートでコントロール一覧取得
- [ ] 各種コントロールタイプの認識確認
- [ ] 各セクションのコントロール取得確認
- [ ] サブフォーム/サブレポートの扱い確認

### 7.2 プロパティテスト

- [ ] 主要プロパティの読み取り確認
- [ ] 各データ型の変換確認
- [ ] 読み取り専用プロパティの扱い確認
- [ ] 存在しないプロパティのエラー確認

### 7.3 書き込みテスト

- [ ] 各種プロパティの設定確認
- [ ] 変更の永続化確認
- [ ] 無効な値のエラー確認
- [ ] 同時アクセス時のエラー確認

### 7.4 エッジケース

- [ ] 日本語コントロール名の処理
- [ ] 特殊文字を含むプロパティ値
- [ ] 非常に長い値の処理
- [ ] Null値の処理

---

## 付録A: 主要コントロールの共通プロパティ

### A.1 全コントロール共通

| プロパティ | 型 | 説明 |
|------------|------|------|
| Name | String | コントロール名 |
| ControlType | Integer | コントロールタイプID |
| Section | Integer | 配置セクションID |
| Left | Long | 左位置（twip単位） |
| Top | Long | 上位置（twip単位） |
| Width | Long | 幅（twip単位） |
| Height | Long | 高さ（twip単位） |
| Visible | Boolean | 表示/非表示 |
| Tag | String | タグ（ユーザー定義データ） |

※ 1 inch = 1440 twips, 1 cm ≈ 567 twips

### A.2 データ連結コントロール（TextBox, ComboBox, ListBox等）

| プロパティ | 型 | 説明 |
|------------|------|------|
| ControlSource | String | データソース（フィールド名または式） |
| DefaultValue | String | 既定値 |
| ValidationRule | String | 入力規則 |
| ValidationText | String | エラーメッセージ |
| Enabled | Boolean | 有効/無効 |
| Locked | Boolean | ロック状態 |
| TabStop | Boolean | タブストップ |
| TabIndex | Integer | タブ順序 |

### A.3 外観関連

| プロパティ | 型 | 説明 |
|------------|------|------|
| BackColor | Long | 背景色 |
| BackStyle | Integer | 背景スタイル（0:透明, 1:標準） |
| BorderColor | Long | 境界線色 |
| BorderStyle | Integer | 境界線スタイル |
| BorderWidth | Integer | 境界線幅 |
| ForeColor | Long | 前景色 |
| FontName | String | フォント名 |
| FontSize | Integer | フォントサイズ |
| FontWeight | Integer | フォントの太さ |
| FontItalic | Boolean | 斜体 |
| FontUnderline | Boolean | 下線 |

---

## 付録B: 色の扱い

Accessの色はLong整数で表現される（OLE Color形式）。

### RGB値からの変換

```
OLE Color = R + (G * 256) + (B * 65536)
```

### 主要な色値

| 色 | 値 |
|----|-----|
| 白 | 16777215 |
| 黒 | 0 |
| 赤 | 255 |
| 緑 | 65280 |
| 青 | 16711680 |
| 黄 | 65535 |
| システムボタン面 | -2147483633 |

※ 負の値はシステムカラーを示す

---

## 更新履歴

| 日付 | バージョン | 内容 |
|------|------------|------|
| 2026-01-01 | 1.0 | 初版作成 |
