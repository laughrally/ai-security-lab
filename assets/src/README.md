# assets/src — 受講画面のスクリプトの正本

`assets/app.mjs` は、このフォルダの `*.part.js` を番号の順につないだものです（中身はそのまま）。
直すときはここを直し、`python3 ~/asl-worker/course/course-sync.py` で app.mjs を作り直してください。
`bash ~/asl-worker/course/check-all.sh` は、app.mjs が src と食い違っていると止まります。

| ファイル | 内容 |
|---|---|
| 00-core | 講座の構成（COURSE ブロック。course.json から自動）・共通の関数（escHtml・csvCell・store） |
| 01-apply-form | 個人の申込フォームと決済画面への遷移 |
| 02-theme | ダークモード |
| 03-auth | Supabase の初期化・ログイン・プラン取得・端末の受講データの持ち主 |
| 04-tabs-quiz-time-access | 理解度チェック・受講時間の実測・タブを開けるかの判定（tabDenyReason）・プラン案内 |
| 05-legal-downloads | 特商法などの表示・ガイドラインのひな形のダウンロード |
| 06-admin-logs-mfa | 管理者：受講ログの出力・二段階認証 |
| 07-corporate-progress | 法人の担当者：受講者の進捗 |
| 08-chapter-work-test | 章末ワーク・章末確認テスト |
| 09-completion-certificate | 受講済みの管理・修了証の申請 |
| 10-exam-data | 認定試験の設問（正解はサーバーだけ） |
| 11-records-report | 受講実績レポート |
| 12-dashboard | 進捗ダッシュボード |
| 13-learn-ui-shell | 学習画面の枠・サイドバーのタブ |
| 14-podcast | 序章の音声プレーヤー |
| 15-learn-ui-lessons | 学習画面：本文の配信・前後ナビ・目次・最初の一歩・受講済みボタン |
| 16-exam-ui | 認定試験の画面 |
| 17-corporate-apply | 法人一括申込 |
| 18-admin-applications | 管理者：法人申込の一覧と請求書 |

すべてのファイルは1つのモジュールとしてつながるので、上のファイルで定義したものを下のファイルで使えます。
ページの onclick から呼ぶ関数は `window.xxx = xxx` で公開しています。
