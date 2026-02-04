# ネットワーク図作成ツール - WPFアプリ

## 📋 必要な環境

- **Visual Studio 2022** (Community Edition以上)
- **.NET 8.0 SDK** (Visual Studioに含まれます)
- **Windows 10/11**

## 🚀 ビルド手順

### 方法1: Visual Studioを使用

1. **Visual Studio 2022を起動**

2. **プロジェクトを開く**
   - 「ファイル」→「開く」→「プロジェクト/ソリューション」
   - `NetworkDiagramApp.csproj` を選択

3. **NuGetパッケージの復元**
   - Visual Studioが自動的に復元します
   - または、「ツール」→「NuGetパッケージマネージャー」→「ソリューションのNuGetパッケージの管理」

4. **ビルド**
   - メニュー: 「ビルド」→「ソリューションのビルド」
   - またはキーボード: `Ctrl + Shift + B`

5. **実行**
   - メニュー: 「デバッグ」→「デバッグなしで開始」
   - またはキーボード: `Ctrl + F5`

### 方法2: コマンドラインを使用

```powershell
# プロジェクトディレクトリに移動
cd NetworkDiagramApp

# ビルド
dotnet build -c Release

# 実行
dotnet run -c Release
```

## 📦 EXEファイルの作成

### シングルファイルEXEの作成

```powershell
cd NetworkDiagramApp

# シングルファイルEXEを生成（.NET Runtimeを含む）
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true

# 出力先: bin\Release\net8.0-windows\win-x64\publish\
```

生成されたEXEファイルは **他のPCでも動作します**（.NET未インストールでもOK）

### フレームワーク依存EXE（サイズ小）

```powershell
# .NET Runtimeがインストール済みの環境向け
dotnet publish -c Release -r win-x64 --self-contained false -p:PublishSingleFile=true

# 出力先: bin\Release\net8.0-windows\win-x64\publish\
```

## 📂 ファイル構成

```
NetworkDiagramApp/
├── NetworkDiagramApp.csproj          # プロジェクトファイル
├── App.xaml                          # アプリケーション定義
├── App.xaml.cs                       # アプリケーションコード
├── MainWindow.xaml                   # メインウィンドウUI
├── MainWindow.xaml.cs                # メインウィンドウロジック
├── Models/
│   └── NetworkNode.cs                # データモデル
└── Services/
    ├── ExcelReader.cs                # Excel読み込み
    ├── TreeBuilder.cs                # ツリー構築
    ├── TreeLayoutCalculator.cs       # レイアウト計算
    ├── VLANFrameManager.cs           # VLAN枠管理
    └── DrawIOGenerator.cs            # Draw.io生成
```

## 🎨 使い方

1. **アプリを起動**
2. **「📂 参照」ボタンでExcelファイルを選択**
3. **出力ファイル名を入力**（省略可）
4. **「🚀 図を生成」ボタンをクリック**
5. **完了ダイアログで「はい」を選択するとフォルダが開きます**

## ⚙️ カスタマイズ

### レイアウト設定の変更

`TreeLayoutCalculator.cs` で以下の定数を編集：

```csharp
private const int NODE_WIDTH = 120;    // ノードの幅
private const int NODE_HEIGHT = 70;    // ノードの高さ
private const int GAP_X = 30;          // 横の間隔
private const int GAP_Y = 80;          // 縦の間隔
```

### VLAN色の変更

`DrawIOGenerator.cs` で色配列を編集：

```csharp
private static readonly string[] VLAN_COLORS = new[]
{
    "#dae8fc",  // 青系
    "#d5e8d4",  // 緑系
    "#fff2cc",  // 黄系
    // ...
};
```

## 🐛 トラブルシューティング

### ビルドエラー: "SDK not found"

→ .NET 8.0 SDKをインストール
https://dotnet.microsoft.com/download/dotnet/8.0

### 実行エラー: "Excel読み込み失敗"

→ Excelファイルを閉じてから実行

### NuGetエラー

```powershell
dotnet restore
```

## 📝 ライセンス

このツールはMITライセンスです。自由に使用・改変できます。

## 🔧 技術スタック

- **.NET 8.0**
- **WPF (Windows Presentation Foundation)**
- **DocumentFormat.OpenXml** - Excel読み込み
- **MVVM パターン** - データバインディング

## 📞 サポート

問題が発生した場合は、エラーメッセージのスクリーンショットと共にお知らせください。
