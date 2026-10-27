// ============================================================
// ★ ここだけ書き換えます(手順は SETUP.md)
// ============================================================
self.APP_CONFIG = {
  // Apps Script の「ウェブアプリ」のURL(デプロイした /exec で終わるもの)
  gasUrl: 'https://script.google.com/macros/s/ここにデプロイIDを入れる/exec',

  // アプリの名前(画面のタイトル・通知に使う)
  appName: 'juki2023 base',

  // Firebase コンソール → プロジェクトの設定 → 「マイアプリ」のウェブアプリの firebaseConfig の中身
  firebase: {
    apiKey: '',
    authDomain: '',
    projectId: '',
    storageBucket: '',
    messagingSenderId: '',
    appId: ''
  },

  // Firebase コンソール → プロジェクトの設定 → Cloud Messaging → ウェブプッシュ証明書 の「キーペア」
  vapidKey: ''
};
