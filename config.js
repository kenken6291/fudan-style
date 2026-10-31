// ============================================================
// FUDAN, 設定ファイル
//   GASを再デプロイしてURLが変わったら API_URL を書き換えてください。
//   リリースのたびに APP_VERSION / RELEASE_DATE / CHANGELOG を更新します。
// ============================================================
window.FUDAN_CONFIG = {
  // GAS ウェブアプリのURL
  API_URL: 'https://script.google.com/macros/s/AKfycbzOgu1bApG9tSIylfsxJjEc1ykIx8irvlhYrCU87tT7xCOAfQz5oJ0ZlaUd4MYBPq4/exec',

  // フロントエンドのバージョン情報
  APP_VERSION: '2026.10.31',
  RELEASE_DATE: '2026-10-31',

  // 更新履歴（新しい順）
  CHANGELOG: [
    { version: '2026.10.31', date: '2026-10-31', notes: [
      '試着イメージをGeminiでも作成できるようにしました（LightXと選択可）',
      '試着イメージの作り直しに対応',
      'GAS URLを設定ファイル（config.js）に分離',
      'バージョン情報の表示を追加'
    ]}
  ]
};
