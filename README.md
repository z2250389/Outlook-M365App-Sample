# Outlook M365 App Minimal Starter

Outlook の左サイドバーに配置する、最小の Microsoft 365 統合アプリ スターターです。

このスターターは次だけに絞っています。

- 左サイドバーのアプリアイコンから個人タブを開く
- タブの読み込み直後に指定 URL を外部ブラウザで開く
- できるだけ新しいブラウザウィンドウで表示する

余計な機能は入れていません。

## 想定構成

- `web/`: サイドバーから開かれる最小ページ
- `manifest/`: Microsoft 365 unified manifest テンプレートと生成物
- `scripts/`: manifest 生成と ZIP パッケージ化
- `assets/`: 配布パッケージ用アイコン

## 動作概要

1. Outlook の左サイドバーでアプリを開きます。
2. タブの読み込み直後に `window.open()` で対象 URL を開きます。
3. ポップアップが許可されていれば、新しいブラウザウィンドウを優先して開きます。
4. ブロックされた場合だけ、タブ内にフォールバックの「開く」ボタンを表示します。

## 既定値

- アプリ名: `サイドバーランチャー`
- 開発元名: `ITOCHU Techno-Solutions Corp.`
- 起動 URL: `https://outlook.office.com/mail/`
- コンテンツ配信元: `https://raw.githack.com/<owner>/<repo>/<branch>/web`

このリポジトリでは、`origin` の GitHub URL と現在のブランチ名から配信元 URL を自動生成します。

## 使い方

1. この公開 GitHub リポジトリへコミットして push します。
2. `scripts/New-M365StarterManifest.ps1` を実行します。
3. `scripts/New-M365StarterPackage.ps1` を実行して配布用 ZIP を作ります。
4. 生成した ZIP を Outlook / Microsoft 365 の管理画面からアップロードします。

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

manifest だけ生成したい場合:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\New-M365StarterManifest.ps1 `
  -OutputPath .\manifest\generated\manifest.json
```

## 注意

- 外部サイトを Outlook 内に iframe 表示する構成ではありません。
- Entra 認証を挟む URL や `X-Frame-Options` / `frame-ancestors` で埋め込みが禁止される URL でも、外部ブラウザ起動なら扱いやすくなります。
- 新しいウィンドウで開くか、新しいタブで開くかは、最終的にはブラウザ側の設定やポップアップ制御の影響を受けます。
- `manifest.json` の `validDomains` には、配信先ドメインと表示対象 URL のドメインを必ず含めてください。
- 既定の配信元は `raw.githack.com` なので、manifest の `validDomains` にそのドメインが入ります。
