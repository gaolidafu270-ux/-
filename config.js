// ============================================================
// ★ ここだけ書き換えます(手順は SETUP.md)
// ============================================================
self.APP_CONFIG = {
  // Apps Script の「ウェブアプリ」のURL(デプロイした /exec で終わるもの)
  gasUrl: 'https://script.google.com/a/macros/city.higashiosaka-osk.ed.jp/s/AKfycbz-BtbObzXHkEUtW2d2ExizzGmwbIyvKBXukJ56xjkOKkKHDOSXIHZgR6rttIGZKMMA/exec',

  // 開き方。'redirect' = アプリ本体のページへ直接移動する(アクセスが「全員」にできない学校のアカウント用)
  //        'frame'    = この外枠の中にアプリを表示する(アクセスが「全員」のとき用)
  mode: 'redirect',

  // アプリの名前(画面のタイトル・通知に使う)
  appName: 'juki2023 base',

  // Firebase コンソール → プロジェクトの設定 → 「マイアプリ」のウェブアプリの firebaseConfig の中身
  firebase: {
    apiKey: 'AIzaSyCE7R5ho-34sZRLll3fPhAKROiPy6a4OT0',
    authDomain: 'juki2023-df581.firebaseapp.com',
    projectId: 'juki2023-df581',
    storageBucket: 'juki2023-df581.firebasestorage.app',
    messagingSenderId: '343379242358',
    appId: '1:343379242358:web:01463768c59ad2eee3de57'
  },

  // Firebase コンソール → プロジェクトの設定 → Cloud Messaging → ウェブプッシュ証明書 の「キーペア」
  vapidKey: 'BHf-yZhX_rmAe10CB2iJQUVd8J4OYgSCYBir4s2GX_nOEPWcnBpPuvz2w3CqPhmmydZr5VY0Ik4keW36s_zZ11M'
};
