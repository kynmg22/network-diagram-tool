# 自動更新機能の設定手順

## 📋 前提条件

- GitHubアカウント
- リポジトリの作成

## 🔧 設定手順

### 1. GitHubリポジトリを作成

1. GitHubにログイン
2. 「New repository」をクリック
3. リポジトリ名を入力（例: `network-diagram-tool`）
4. Public または Private を選択
5. 「Create repository」をクリック

### 2. コードを編集

`Services/UpdateChecker.cs` を開いて、以下を編集：

```csharp
// TODO: GitHubのユーザー名とリポジトリ名に置き換えてください
private const string GITHUB_OWNER = "your-username";  // ← あなたのGitHubユーザー名
private const string GITHUB_REPO = "network-diagram-tool";  // ← リポジトリ名
```

### 3. バージョン番号を管理

新しいバージョンをリリースする時は、`UpdateChecker.cs` のバージョン番号を更新：

```csharp
private const string CURRENT_VERSION = "1.1.0";  // ← バージョンを上げる
```

### 4. GitHub Releasesでリリース

1. GitHubのリポジトリページを開く
2. 「Releases」→「Create a new release」
3. **Tag version**: `v1.1.0` のように入力（先頭に `v` をつける）
4. **Release title**: `バージョン 1.1.0`
5. **Description**: 更新内容を記載
   ```
   ## 変更内容
   - シート選択機能を追加
   - アイコンを追加
   - バグ修正
   ```
6. **Attach binaries**: 生成した `.exe` ファイルをドラッグ&ドロップ
7. 「Publish release」をクリック

### 5. 動作確認

アプリを起動すると：
- 自動的に最新バージョンをチェック
- 新しいバージョンがあれば通知が表示される
- 「はい」をクリックすると、ダウンロードページが開く

## 📝 リリースの流れ

```
1. コードを修正
   ↓
2. UpdateChecker.csのバージョン番号を上げる (例: 1.0.0 → 1.1.0)
   ↓
3. EXEをビルド
   dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
   ↓
4. GitHubで新しいReleaseを作成
   - Tag: v1.1.0
   - EXEファイルを添付
   ↓
5. ユーザーがアプリを起動すると自動的に通知される
```

## 🔒 プライベートリポジトリの場合

プライベートリポジトリでも動作しますが、以下の設定が必要です：

1. GitHub Personal Access Token (PAT) を作成
2. `UpdateChecker.cs` に以下を追加：
   ```csharp
   client.DefaultRequestHeaders.Add("Authorization", "Bearer YOUR_TOKEN");
   ```

ただし、**トークンをソースコードに埋め込むのはセキュリティリスク**があるため、パブリックリポジトリの使用を推奨します。

## ❓ トラブルシューティング

### 更新チェックが動作しない

1. **インターネット接続を確認**
2. **GitHubユーザー名とリポジトリ名が正しいか確認**
3. **Releaseが作成されているか確認**
4. **Tag名が `v` で始まっているか確認**（例: `v1.0.0`）

### エラーメッセージ

- `HTTP 404`: リポジトリが見つからない → GITHUB_OWNER と GITHUB_REPO を確認
- `接続がタイムアウト`: ネットワーク接続を確認
- `更新チェック失敗`: Releaseが公開されているか確認

## 🎯 ベストプラクティス

- **セマンティックバージョニング**を使用: `major.minor.patch`
  - `major`: 互換性のない変更（例: 2.0.0）
  - `minor`: 機能追加（例: 1.1.0）
  - `patch`: バグ修正（例: 1.0.1）

- **リリースノート**に必ず更新内容を記載
- **定期的にバージョンアップ**してユーザーに通知
