// ============================================================
// ★ ここだけ書き換えます(手順は SETUP.md)
// ============================================================
self.APP_CONFIG = {AIzaSyCE7R5ho-34sZRLll3fPhAKROiPy6a4OT0
  // Apps Script の「ウェブアプリ」のURL(デプロイした /exec で終わるもの)
  gasUrl: 'https://script.google.com/macros/s/AKfycbwjHAy4U_RLuD0N7-QyQhG9JHfDT0iWrxR8hk8KVCLulMCmxeuEvQs26t7v6hQilFq9/exec',

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
  vapidKey: 'BHf-yZhX_rmAe10CB2iJQUVd8J4OYgSCYBir4s2GX_nOEPWcnBpPuvz2w3CqPhmmydZr5VY0Ik4keW36s_zZ11M'
};
