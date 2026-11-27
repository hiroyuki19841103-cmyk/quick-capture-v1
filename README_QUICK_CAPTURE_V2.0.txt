QUICK CAPTURE V2.0（2026年11月27日）

■ 新規モード「Digital NOLTYメモ」
Weekly：指定日の週間画面下部の日別メモ（DAY）へ追記。
Daily：指定日のDaily画面下部の専用メモ（DAILY_VIEW）へ追記。
WeeklyとDailyは別々のメモ。週間上部の週前半／週後半メモには書き込みません。
従来のメモ／タスクと同じく、オンライン登録またはオフライン一時保存が可能です。
既存のメモがある場合は空行を挟んで追記します。同期成功時だけ未同期から削除します。
通信復帰・自動認証だけでは送信せず、「オンライン保存」から手動実行します。

■ GitHub上書き／追加対象（QUICK CAPTUREのみ）
上書き：src/main.js、src/app-config.js、src/styles.css
新規：src/nolty-sync.js、src/nolty-memo-workbook.js、src/Digital_NOLTY_Memos_EDITABLE_TEMPLATE.xlsx
Digital NOLTY FINAL8.7の既存コード、TASK CONTROL、IndexedDB、Service Workerは変更不要。

■ 初回のみ Microsoft Entra の設定
Digital NOLTYのアプリ登録（クライアントID：d3b2320f-52f9-40a1-b2d7-7ae1707994b5）の「認証」→「シングルページアプリケーション（SPA）」へ次のリダイレクトURIを追加する。
https://black-pond-0a1e56410.2.azurestaticapps.net/
既存のDigital NOLTYのURIは削除しない。TASK CONTROL側のアプリ登録も変更不要。
既存の Files.ReadWrite.AppFolder を利用するため、追加のGraphアクセス許可は不要。
QUICK CAPTUREで通常のMicrosoftログイン後、NOLTY接続を確認。必要なら「NOLTY接続」ボタンで同じMicrosoftアカウントを選択する。

■ 確認手順
1. DIGITAL NOLTY側でOneDrive同期を行い、最新のメモを保存しておく。
2. V2.0起動、NOLTY接続、Weeklyを選んで日付を指定し登録。Digital NOLTYのOneDrive同期後に該当日の週間下部メモへ追記を確認。
3. Dailyも同様に、その日のDaily専用メモへ追記されることを確認。
4. オフラインで両モードを一時保存→再起動して未保存件数を確認→通信復帰→オンライン保存。
5. 通信復帰時は勝手に送信されないこと、同期後に件数が減ること、元メモが残ることを確認。
6. 既存の外部メモ、タスク登録も従来どおり確認する。

■ 注意点
Digital NOLTY側に未同期の編集が残っていると、同じメモの変更で競合確認が表示されることがあります。先にNOLTYの編集を同期し、QUICK CAPTUREから追記後はNOLTYのOneDrive同期で最新内容を取得してください。
Masterが見つからない、形式が異なる、Excelの文字数上限を超える、前回送信結果が不明で他の編集と競合する等の場合は未同期を保持し、勝手に上書きしません。
オフライン保存領域はV1.6.1と共通です。サイトデータを消去しないでください。
