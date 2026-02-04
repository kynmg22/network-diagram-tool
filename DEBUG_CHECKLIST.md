# デバッグチェックリスト - v1.2

## ✅ 修正済み

1. **MainWindow.xaml.cs**
   - ✅ `using System.Linq;` を追加（Contains メソッド用）

2. **MainWindow.xaml**
   - ✅ Grid.RowDefinitions を 9行に増やした（Grid.Row="8" 使用のため）
   - ✅ DropZone（Border）にドラッグ&ドロップイベントを追加
   - ✅ テンプレート作成ボタンを追加

3. **ExcelTemplateGenerator.cs**
   - ✅ 新規作成済み
   - ✅ 必要な using ディレクティブあり

## ⚠️ テスト時の注意点

### 1. ドラッグ&ドロップ機能

**期待される動作:**
```
1. Excelファイルをウィンドウにドラッグ
   → 背景が青色（#E6F4FF）に変化
   
2. ドロップ
   → 背景が白色に戻る
   → ログに「✓ ファイルをドロップ: ファイル名.xlsx」
   → シート一覧を自動取得
   
3. .xlsx/.xls 以外をドロップ
   → エラーダイアログ表示
```

**テストケース:**
- ✓ .xlsx ファイル
- ✓ .xls ファイル
- ✗ .txt ファイル（エラー表示）
- ✗ .pdf ファイル（エラー表示）
- ✓ 複数ファイルを選択してドロップ（最初の1つのみ処理）

### 2. Excelテンプレート作成機能

**期待される動作:**
```
1. 「📄 Excelテンプレートを作成」ボタンをクリック
   → 保存ダイアログが表示
   
2. 保存場所とファイル名を指定
   → Excelファイル作成
   → ログに「✓ テンプレートを作成: ファイル名.xlsx」
   → 確認ダイアログ表示
   
3. 「はい」をクリック
   → エクスプローラーでファイルの場所を開く
```

**生成されるExcelの内容:**
- シート名: 「構成図作成」
- ヘッダー行: A列～G列（接続元ID, ID, 機器名, IPアドレス, VLANID, 備考）
- サンプルデータ: 9行（ONU → HUB → Router → Switch → PC, Printer）

### 3. 既存機能との統合

**確認項目:**
- ✓ 参照ボタンは正常動作するか
- ✓ ドラッグ&ドロップ後に参照ボタンも使えるか
- ✓ シート選択ドロップダウンは正常動作するか
- ✓ 「図を生成」ボタンは正常動作するか

## 🐛 予想される問題と対処

### 問題1: ドラッグ中に背景色が変わらない
**原因:** Border要素ではなく子要素がイベントを受け取っている
**対処:** 問題なし（現在の実装で動作するはず）

### 問題2: ドロップしてもファイルが読み込まれない
**原因:** 
- e.Data.GetDataPresent(DataFormats.FileDrop) が false
- ファイルパスの取得失敗

**対処:** 
デバッグログを追加して確認:
```csharp
AddLog($"DataPresent: {e.Data.GetDataPresent(DataFormats.FileDrop)}");
```

### 問題3: テンプレート作成時にエラー
**原因:**
- DocumentFormat.OpenXml の参照がない
- 保存先のフォルダが存在しない

**対処:**
- .csproj で DocumentFormat.OpenXml 3.0.2 が参照されているか確認
- 保存先フォルダの書き込み権限を確認

### 問題4: 日本語が文字化け
**原因:** UTF-8 BOM の問題
**対処:** 現在の実装で UTF-8 BOMなし設定済み（問題なし）

## 📋 ビルド前チェック

1. **すべてのファイルが保存されているか**
2. **NuGetパッケージが復元されているか**
   ```
   dotnet restore
   ```
3. **クリーンビルド**
   ```
   dotnet clean
   dotnet build
   ```

## 🚀 ビルドコマンド

```powershell
cd C:\Users\ky-nomura\Downloads\NetworkDiagramApp_v1.2

# デバッグビルド（開発中）
dotnet build

# リリースビルド（配布用）
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:PublishReadyToRun=false -p:PublishTrimmed=true
```

## 📝 テスト手順

### Step 1: 基本動作確認
1. アプリ起動
2. ログに起動メッセージが表示されるか
3. すべてのボタンが表示されているか

### Step 2: ドラッグ&ドロップテスト
1. Excelファイルをドラッグ → 背景色変化を確認
2. ドロップ → ファイルパス表示を確認
3. シート一覧が表示されるか確認

### Step 3: テンプレート作成テスト
1. 「テンプレート作成」ボタンクリック
2. ファイルを保存
3. Excelで開いて内容確認

### Step 4: 統合テスト
1. テンプレートを作成
2. 作成したテンプレートをドラッグ&ドロップ
3. シートを選択
4. 図を生成
5. draw.ioで開いて確認

## ✅ すべて正常なら

GitHub にコミット＆プッシュ
→ v1.2.0 としてリリース作成
