QUICK CAPTURE V1｜TASK CONTROL専用入力アプリ
作成日：2026-09-20（日本時間）

【目的】
TASK CONTROLの一覧や期限を見ず、メモまたはタスクだけを同じExcel Masterへ即時登録します。

【動作】
メモ：分類「外部メモ」で登録。Microsoft To Do同期対象外。
タスク：選択した通常分類で登録。TASK CONTROLの次回同期時にMicrosoft To Doへ登録。
登録ボタン：最新Excel Masterを取得し、新しい行を追記してOneDriveへ即時同期。
競合：更新直前にExcelが変わった場合は最新版を再取得して最大4回再試行。
失敗：入力内容を端末内に保持し「未送信を再同期」から再送。
連続入力：登録後も同じモードを維持可能。

【重要】
・新しいExcel Masterは作成しません。TASK CONTROLと同じAppFolder内のTASK_CONTROL_Master.xlsxを使用します。
・初回は先にTASK CONTROLでOneDrive同期し、Excel MasterをAppFolderへ作成してください。
・Excelデスクトップ／WebアプリでMasterを開いている間は閉じてください。
・QUICK CAPTUREはMicrosoft To Doへ直接接続しません。必要権限はFiles.ReadWrite.AppFolderだけです。

【設定】
1. src/app-config.jsを開きます。
2. CLIENT_IDへTASK CONTROLと同じApplication (client) IDを設定します。
3. REDIRECT_URIへQUICK CAPTUREの実URLを設定します。
4. Entraの同じアプリ登録 → 認証 → Single-page applicationへ、その実URLを追加します。
5. 既存権限Files.ReadWrite.AppFolderをそのまま使います。新しいAPI権限は不要です。

【Azure Static Web Appsへの配置】
1. 本フォルダーの中身を新しいGitHubリポジトリのルートへ置きます。
2. Azure Static Web Appsを作成します。
3. Build preset：Custom
4. App location：/
5. Output location：dist
6. デプロイ後のURLをEntraとsrc/app-config.jsの両方へ完全一致で設定します。

【初回確認】
□ Microsoftログインできる
□ メモを登録すると緑色でOneDrive同期完了になる
□ TASK CONTROLで同期すると外部メモとして表示される
□ 外部メモがMicrosoft To Doには存在しない
□ 外部メモを防衛等へ変更するとTo Doへ登録される
□ タスクモードの内容がTASK CONTROLへ反映される
□ PC、iPad、Z Fold7で登録できる

【注意】
同時更新競合は自動再試行しますが、Excel Masterを開いたままのロックは解除できません。その場合も入力内容は端末内に残ります。
