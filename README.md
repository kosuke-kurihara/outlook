# English Reply Headers — 静的ホスト版 1.3

新しいOutlook / Outlook on the webで、日本語の返信・転送ヘッダーを英語に変換する試作アドインです。Outlookの表示言語は変更しません。通常の利用にNode・npm・localhostサーバーは不要です。

## 現在の状態

旧localhost版をもとに、常駐サーバー不要の静的ホスト版へ移行しました。

手動ボタン用と管理者配布用のテンプレート、オフラインの設定ページ、変換処理、テストを同梱しています。公開先は https://kosuke-kurihara.github.io/outlook/ です。設定済みマニフェストは docs/manifest-user.xml（手動版）と docs/manifest-admin-auto.xml（管理者配布版）です。GitHub Pagesの有効化が必要です。

ローカルのMicrosoft Edgeで変換テストと設定ページの生成テストは成功しています。2026-09-11にMicrosoft公式 office-addin-manifest 3.0.0 で両版のマニフェストが有効と判定されました。Outlook実機での検証は未実施です。実メールで使う前に、下記の検証を行ってください。

## 導入

1. このフォルダーの内容を使用するGitHubリポジトリへ配置します。
2. GitHub Pagesの公開元をmainブランチの `/docs` に設定します（別のHTTPS静的ホストでも利用可能）。
3. 公開URLと、その下の `runtime.html`、`runtime.js`、`assets/icon-32.png` が取得できることを確認します。
4. ローカルの `configure.html` をブラウザーで開き、公開URL（例 `https://USERNAME.github.io/REPO`）を入力します。`/docs` は公開URLに含めません。
5. 「手動ボタン（個人用）」を選択し `manifest-user.xml` を保存します。
6. 開発環境がある場合は `npx office-addin-manifest validate manifest-user.xml` で公式ツールの検証を行います。これは開発時のみ必要です。
7. Outlookのカスタムアドイン管理画面からマニフェストを読み込みます。
8. テスト用メールの返信・転送画面で「English headers」を押します。

旧版と同じアドインIDを引き継いでいます。手動版と自動版も同じIDであり、併用する二つのアドインではありません。切り替え時は置き換えとして扱います。

## 変換対象

- HTMLは最初の `divRplyFwdMsg` または `x_divRplyFwdMsg` / `divRplyFwdMsg_1` 等のブロックを対象とします。
- `b` / `strong` 内に From、Sent、To、任意のCc、Subject相当の見出しが順に存在する場合のみ変換します。
- すでに英語の最新ヘッダーがあれば終了し、古い日本語の引用へ進みません。
- プレーンテキストも見出しの連続した並びを確認します。折り返された複数行の宛先や異なるマークアップは、保守的に変換しない場合があります。
- 日付は文字列の書式だけを変更します。タイムゾーンは変換しません。
- `docs/runtime.js` の `addOriginalMessageSeparator` を `false` にすると区切り行の挿入を無効化できます。

## 本文更新と制約

`ReadWriteItem` 権限を使い、`getAsync` / `setAsync` に同じ `BodyMode.HostConfig` を渡します。Outlookの会話表示設定によって対象本文の範囲が変わり、最新返信部分のみ取得されて引用ヘッダーがない場合は何も変更しません。

本文更新前に再読み取りし、内容が変わっていた場合は書き込みを中止します。ただしこれは原子的な比較更新ではないため、その後の短い時間に発生する編集との競合は完全には防げません。ボタン操作中は入力を止め、変換後に本文を確認してください。

Office.jsに任意のヘッダーブロックだけを直接置換するAPIがないため、更新には本文の `setAsync` を使います。カーソル位置・書式・署名・引用の折りたたみ・インライン画像への影響は実機での確認が必要です。HTML全体のパース・再シリアライズも行うため、対象外のHTMLがバイト単位で一致する保証はありません。

自動版は `OnNewMessageCompose`（Mailbox 1.10）を使用する実験版です。新規メールはスキップし、返信・転送で処理します。既存下書きの再編集はイベント対象外です。自動実行には管理者配布が必要です。HTML処理にはDOMParserを使い、使用できないランタイムでは変更しません。Classic Outlookは対象外です。送信を止めるイベントは使いません。

本文の外部送信・保存・ログ出力は実装していません。HTTPSホストへのアセット取得とMicrosoftのOffice.js読み込みは必要です。実メールをGitHubへアップロードしないでください。

## 検証

開発時のみNodeとPlaywright、およびMicrosoft Edgeを使用します。

```text
npm install --no-save playwright
node tests/test.cjs
```

2026-09-11の実行では32件のチェックが成功しました（同じAPIオプション確認の繰り返しを含む）。日付、誤検出、引用チェーン、入れ子、署名保持、再実行、編集中の変更、XML生成、URL拒否などを確認しています。合成データとOffice.jsモックを使ったテストで、Outlookの実機テストではありません。

実機では送信せず、テスト下書きで以下を確認します。

- 新しいOutlookとWeb版それぞれの返信・全員に返信・転送
- HTML / プレーンテキスト
- 長い引用チェーン、署名、リンク、インライン画像
- 会話表示と個別表示、ポップアウト
- ボタンを2回押して古い引用や区切り行が変化しないこと
- 変換後に本文と宛先を目視確認

実機のヘッダーが認識されない場合は、匿名化した `body.getAsync` のHTML例をもとにパーサーを調整する必要があります。

## 仕様確認の出典

- [本文の取得と更新 / bodyMode](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/insert-data-in-the-body)
- [Outlookアドインの権限](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions)
- [Body API / setAsyncの制約](https://learn.microsoft.com/en-us/javascript/api/outlook/office.body?view=outlook-js-preview)
- [イベントと管理者配布](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/event-based-activation)
- [getComposeTypeAsync](https://learn.microsoft.com/en-us/javascript/api/outlook/office.messagecompose?view=outlook-js-preview)
