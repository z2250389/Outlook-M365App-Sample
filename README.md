# Outlook M365 App Minimal Starter

Outlook の左サイドバーに配置する最小の Microsoft 365 統合アプリ スターターです。

このスターターは次だけに絞っています。

- 左サイドバーのアプリアイコンから個人タブを開く
- タブの読み込み直後に URL ベースのダイアログを開く
- ダイアログ内で指定 URL を WebView 表示する

余計な機能は入れていません。

## 想定構成

- `web/`: サイドバーから開かれる最小ページとダイアログページ
- `manifest/`: Microsoft 365 unified manifest テンプレート
- `scripts/`: manifest 生成と ZIP パッケージ化
- `assets/`: 配布パッケージ用アイコン

## 使い方

1. この公開 GitHub リポジトリへコミットして push します。
2. `scripts/New-M365StarterManifest.ps1` を実行します。
3. `scripts/New-M365StarterPackage.ps1` を実行して配布用 ZIP を作ります。
4. 生成した ZIP を Outlook / Microsoft 365 の管理経由で配布します。

既定値:

- アプリ名: `サイドバーランチャー`
- 起動 URL: `https://www.ctc-g.co.jp/`
- コンテンツ配信元: `https://cdn.jsdelivr.net/gh/<owner>/<repo>@<branch>/web`

このリポジトリでは、`origin` の GitHub URL と現在のブランチ名から配信元 URL を自動生成します。

例:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\New-M365StarterPackage.ps1
```

値を上書きしたい場合:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\New-M365StarterPackage.ps1 `
  -TargetUrl "https://example.com" `
  -AppName "別タイトル"
```

## 注意

- URL ベースの `dialog` API は TeamsJS 上で提供されます。
- `dialog.url` は Microsoft Learn 上では preview 扱いです。実運用前に対象の Outlook / Microsoft 365 クライアントで確認してください。
- `cdn.jsdelivr.net` を配信元に使うため、manifest の `validDomains` にそのドメインが入ります。
- 表示先 URL が iframe 埋め込みを禁止している場合、ダイアログ内表示はできません。その場合はブラウザー遷移リンクを使ってください。
- `manifest.json` の `validDomains` には、配信先ドメインと表示対象 URL のドメインを必ず含めてください。
