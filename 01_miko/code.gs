// ============================================================
//  main.js  ―  みこちゃん LINE × Dify Bot 完全版
//  Google Apps Script
//
//  スプレッドシート列構成:
//    A: userID
//    B: conversationId
//    C: status (free / premium)
//    D: 最終利用日
//    E: 当日利用回数
//    F: 登録日
//    G: 気分スコア履歴
//    H: stripeCustomerId
//
//  スクリプトプロパティ:
//    LINE_ACCESS_TOKEN / LINE_CHANNEL_SECRET
//    DIFY_API_KEY / SPREADSHEET_ID
//    STRIPE_SECRET_KEY / STRIPE_BASE_LINK
//    STRIPE_WEBHOOK_SECRET / ADMIN_LINE_USER_ID
// ============================================================

var scriptProperties      = PropertiesService.getScriptProperties();
var LINE_ACCESS_TOKEN     = scriptProperties.getProperty('LINE_ACCESS_TOKEN');
var LINE_CHANNEL_SECRET   = scriptProperties.getProperty('LINE_CHANNEL_SECRET');
var DIFY_API_KEY          = scriptProperties.getProperty('DIFY_API_KEY');
var DIFY_API_URL          = 'https://api.dify.ai/v1/chat-messages';
var STRIPE_SECRET_KEY     = scriptProperties.getProperty('STRIPE_SECRET_KEY');
var STRIPE_WEBHOOK_SECRET = scriptProperties.getProperty('STRIPE_WEBHOOK_SECRET');
var ADMIN_LINE_USER_ID    = scriptProperties.getProperty('ADMIN_LINE_USER_ID');
var SHEET_NAME            = 'users';
var FREE_DAILY_LIMIT      = 5;

// ============================================================
//  エントリポイント
// ============================================================
function doPost(e) {
  if (!e || !e.postData) return okResponse();

  // Stripe Webhookの処理
  var stripeSig = (e.parameter && e.parameter['stripe-signature'])
               || (e.headers   && e.headers['Stripe-Signature']);
  if (stripeSig) {
    return handleStripeWebhook(e.postData.contents, stripeSig);
  }

  try {
    var body          = e.postData.contents;
    var lineSignature = e.headers && e.headers['X-Line-Signature'];

    if (!verifyLineSignature(body, lineSignature)) {
      Logger.log('LINE署名検証失敗');
      return okResponse();
    }

    var postData = JSON.parse(body);
    if (!postData.events || postData.events.length === 0) return okResponse();

    var event = postData.events[0];

    // 重複メッセージ防止
    if (event.message && event.message.id) {
      var msgId = event.message.id;
      var cache = CacheService.getScriptCache();
      if (cache.get('msg_' + msgId)) {
        Logger.log('重複メッセージをスキップ: ' + msgId);
        return okResponse();
      }
      cache.put('msg_' + msgId, '1', 60);
    }

    if (event.type === 'message' && event.message && event.message.text) {
      handleTextMessage(event);
    }
  } catch (err) {
    Logger.log('doPost error: ' + err.message);
  }
  return okResponse();
}

// ============================================================
//  LINE署名検証
// ============================================================
function verifyLineSignature(body, signature) {
  if (!LINE_CHANNEL_SECRET || !signature) {
    Logger.log('署名検証スキップ（未設定）');
    return true;
  }
  try {
    var hash     = Utilities.computeHmacSha256Signature(
      body, LINE_CHANNEL_SECRET, Utilities.Charset.UTF_8
    );
    var expected = Utilities.base64Encode(hash);
    return expected === signature;
  } catch (err) {
    Logger.log('署名検証エラー: ' + err.message);
    return false;
  }
}

// ============================================================
//  メッセージ処理（みこ語尾統一版）
// ============================================================
function handleTextMessage(event) {
  var userId      = event.source.userId;
  var userMessage = event.message.text.trim();
  var replyToken  = event.replyToken;

  showLoadingAnimation(userId);

  var sheet   = getSheet();
  var today   = getTodayString();
  var userRow = getUserRow(sheet, userId);

  // 新規ユーザー登録
  if (!userRow) {
    sheet.appendRow([userId, '', 'free', today, 0, today, '', '']);
    userRow = sheet.getLastRow();
    notifyAdmin('新規ユーザーが登録したよ！\nUserID: ' + userId);
  }

  // 日付リセット処理
  var lastDateRaw = sheet.getRange(userRow, 4).getValue();
  var lastDate    = formatDateValue(lastDateRaw);
  var count       = Number(sheet.getRange(userRow, 5).getValue()) || 0;
  if (lastDate !== today) {
    count = 0;
    sheet.getRange(userRow, 4).setValue(today);
    sheet.getRange(userRow, 5).setValue(0);
  }

  var status    = sheet.getRange(userRow, 3).getValue();
  var isPremium = (status === 'premium');

  // ============================================================
  // 【最優先】数字1〜5は制限なしで必ず処理
  // ============================================================
  if (userMessage === '1' || userMessage === '2' || userMessage === '3' ||
      userMessage === '4' || userMessage === '5') {
    handleMoodTracking(sheet, userRow, userId, userMessage, replyToken);
    return;
  }

  // ============================================================
  // 【優先】固定コマンド（制限なし）
  // ============================================================

  // 気分を記録ボタン
  if (userMessage === '気分を記録') {
    sendLineReply(replyToken,
      '今の気分を数字で教えてほしいみこ🌸\n\n' +
      '😢 1 → めちゃくちゃしんどい\n' +
      '😔 2 → ちょっとしんどい\n' +
      '😐 3 → まあ普通\n' +
      '🙂 4 → 割といい感じ\n' +
      '😊 5 → 最高！\n\n' +
      '1〜5の数字をポチッと送ってみこ！');
    return;
  }

  // みこに話すボタン
  if (userMessage === 'こんにちは' || userMessage === 'みこに話す') {
    sendLineReply(replyToken,
      'なーに？みこだよ〜！🌸\n\nなんでも話しかけてほしいみこ！\n今日はどんな一日だったみこ？✨');
    return;
  }

  // SOS
  if (userMessage === 'SOS' || userMessage === 'sos') {
    notifyAdmin('⚠️ SOSが送信されたよ。\nUserID: ' + userId);
    sendLineReply(replyToken,
      '…ねえ、大丈夫？\n' +
      'みこちゃん、今すぐ心配してるよ。\n\n' +
      '消えたいとか、傷つけたいって気持ちがあるなら、\n' +
      '一人で抱え込まないでほしいな。\n\n' +
      'みこちゃんがずっとそばにいたいけど、\n' +
      'もっと力になってくれる人たちも紹介させてね。\n\n' +
      '📞 よりそいホットライン\n' +
      '　　0120-279-338（24時間）\n\n' +
      '📞 いのちの電話\n' +
      '　　0120-783-556\n\n' +
      '肩の力、少しだけ抜いてみてね。\n' +
      'みこちゃんはずっとここにいるよ。');
    return;
  }

  // ステータス確認
  if (userMessage === 'ステータス確認') {
    if (isPremium) {
      sendLineReply(replyToken,
        'プレミアムみこ！ありがとうみこ〜！🎉\n\n' +
        '残り回数：無制限みこ！\n\n' +
        'いつでも話しかけてくれて大丈夫みこ！');
    } else {
      sendLineReply(replyToken,
        '今のプランを確認したみこ！\n\n' +
        'プラン：無料みこ\n' +
        '今日の残り：' + (FREE_DAILY_LIMIT - count) + '回みこ\n\n' +
        'もっと話したいなら「プレミアムみこ」あるみこ〜！');
    }
    return;
  }


  // リセット
  if (userMessage === 'リセット' || userMessage === 'reset') {
    sheet.getRange(userRow, 2).setValue('');
    sendLineReply(replyToken,
      'リセット完了みこ〜！🌸\n\n' +
      '気分新たに、また話しかけてほしいみこ！');
    return;
  }

  // 解約処理
  if (userMessage.includes('解約') || userMessage === '解約したい') {
    if (!isPremium) {
      sendLineReply(replyToken,
        '今は無料プランみこ！\n解約の必要はないみこ〜🌸\n\nいつでも話しかけてほしいみこ！');
    } else {
      sendLineReply(replyToken,
        '解約のご連絡ありがとうみこ😢\n\n' +
        '次の更新日までは引き続き使えるみこ！\n\n' +
        'いつでも戻ってきてほしいみこ〜🍓\n' +
        'みこちゃんはずっと待ってるみこ！');
    }
    return;
  }

  // プレミアム申し込み
  if (userMessage.includes('プレミアム') || userMessage.includes('申し込み')) {
    var baseLink = scriptProperties.getProperty('STRIPE_BASE_LINK');
    if (baseLink) {
      var url = baseLink + '?client_reference_id=' + userId;
      sendLineReply(replyToken,
        'プレミアムみこの申し込みページを用意したみこ！🎉\n\n' +
        '✨ 特典：\n' +
        '・無制限でみこちゃんとお話できるみこ！\n' +
        '・毎朝7時「おはようみこ〜！」\n' +
        '・毎昼12時「お昼だみこ〜！」\n' +
        '・毎夜21時「おやすみみこ〜！」\n' +
        '・気分トラッキング機能\n\n' +
        '💳 月額500円（税込）\n\n' +
        '🔓 解約方法：\n' +
        'みこちゃんに「解約したい」と送るだけみこ！\n' +
        'いつでも解約できるみこ🍓\n\n' +
        '↓ 申し込みはこちらみこ〜🌸\n' + url);
    } else {
      sendLineReply(replyToken,
        'ごめんみこ〜！今ページの準備ができてないみこ😢\n' +
        'もう少し待ってほしいみこ！🍓');
    }
    return;
  }

  
  // 地域設定
  const cityMapForSetting = {
    '東京': 'Tokyo', '大阪': 'Osaka', '名古屋': 'Nagoya',
    '福岡': 'Fukuoka', '札幌': 'Sapporo', '仙台': 'Sendai',
    '広島': 'Hiroshima', '京都': 'Kyoto', '神戸': 'Kobe',
    '埼玉': 'Saitama', '横浜': 'Yokohama', '千葉': 'Chiba',
  };
  if (cityMapForSetting[userMessage]) {
    var cityName = userMessage;
    sheet.getRange(userRow, 9).setValue(cityMapForSetting[cityName]);
    sheet.getRange(userRow, 10).setValue(cityName);
    sendLineReply(replyToken,
      cityName + 'に設定したみこ！🌸\n' +
      '毎朝' + cityName + 'の天気をお届けするみこ〜！✨');
    return;
  }


  // ============================================================
  // 【利用制限チェック】Dify API呼び出し前に確認
  // ============================================================
  if (!isPremium && count >= FREE_DAILY_LIMIT) {
    sendLineReply(replyToken,
      'えっみこ！？今日だけでそんなに会いに来てくれたの！？\n' +
      'みこちゃん、うれしすぎるみこ…\n\n' +
      'もっと話したいなら「プレミアムみこ」あるみこ〜！\n' +
      '明日になったらまた話せるみこ！');
    return;
  }

  // ---- Dify API ----
  var conversationId = sheet.getRange(userRow, 2).getValue() || '';
  var result = callDifyApi(userId, userMessage, conversationId);

  if (result.conversationId) sheet.getRange(userRow, 2).setValue(result.conversationId);
  sheet.getRange(userRow, 5).setValue(count + 1);
  sendLineReply(replyToken, result.answer);
}

// ============================================================
//  気分トラッキング（みこ語尾統一）
// ============================================================
function handleMoodTracking(sheet, userRow, userId, score, replyToken) {
  var moodLabels = {
    '1': 'めちゃくちゃしんどい',
    '2': 'ちょっとしんどい',
    '3': 'まあ普通',
    '4': '割といい感じ',
    '5': '最高！'
  };
  var moodEmoji = { '1':'😢', '2':'😔', '3':'😐', '4':'🙂', '5':'😊' };
  var moodFollow = {
    '1': 'すごくしんどいんみこね。無理しなくていいみこ。今どんな気持ちか話してほしいみこ。',
    '2': 'ちょっとしんどいんみこね。何かあれば話しかけてほしいみこ！',
    '3': 'まあ普通かあみこ〜！何かあれば話しかけてほしいみこ！',
    '4': 'いい感じじゃんみこ！何かあれば話しかけてほしいみこ！',
    '5': 'やったみこ〜！最高みこ！！その調子で今日も過ごしてほしいみこ〜！'
  };

  // 履歴を保存
  var historyRaw = String(sheet.getRange(userRow, 7).getValue() || '');
  var history    = historyRaw ? historyRaw.split(',') : [];
  history.push(score);
  if (history.length > 5) history.shift();
  sheet.getRange(userRow, 7).setValue(history.join(','));

  // グラフ表示
  var avg = (history.reduce(function(s, v) { return s + Number(v); }, 0) / history.length).toFixed(1);
  var bar = history.map(function(v) {
    return '●'.repeat(Number(v)) + '○'.repeat(5 - Number(v));
  }).join('\n');

  sendLineReply(replyToken,
    moodEmoji[score] + ' ' + moodLabels[score] + '（' + score + '/5）\n' +
    '記録したみこ！！\n\n' +
    '─ 最近の気分 ─\n' +
    bar + '\n' +
    '平均：' + avg + ' / 5\n\n' +
    moodFollow[score]);
}

// ============================================================
//  Stripe Webhook処理
// ============================================================
function handleStripeWebhook(payload, signature) {
  try {
    var data = JSON.parse(payload);
    Logger.log('Stripe type: ' + data.type);

    // プレミアム契約完了
    if (data.type === 'checkout.session.completed') {
      var session    = data.data.object;
      var lineUserId = session.client_reference_id
                    || (session.metadata && session.metadata.lineUserId);
      if (lineUserId) {
        updateStatus(lineUserId, 'premium');
        saveCustomerMapping(session.customer, lineUserId);
        pushMessage(lineUserId,
          'プレミアムみこ・ようこそみこ〜！🎉\n\n' +
          'ありがとうみこ！これからは制限なしで話せるみこ！\n' +
          'いつでも気軽に話しかけてほしいみこ〜！');
        notifyAdmin('プレミアム契約完了！\nUserID: ' + lineUserId);
      } else {
        notifyAdmin('決済完了したけどLINEユーザーIDが取得できなかったよ。手動でpremiumに更新してね。');
      }
    }

    // 解約処理
    if (data.type === 'customer.subscription.deleted') {
      var customerId = data.data.object.customer;
      var lineUserIdFromCancel = getLineUserIdByCustomer(customerId);
      if (lineUserIdFromCancel) {
        updateStatus(lineUserIdFromCancel, 'free');
        pushMessage(lineUserIdFromCancel,
          'プレミアムプランが終了したんみこね。\n\n' +
          'また話しかけてくれると嬉しいみこ！\n' +
          '無料でも1日5回まで話せるみこ！');
        notifyAdmin('解約発生！\nUserID: ' + lineUserIdFromCancel);
      }
    }
  } catch (err) {
    Logger.log('Stripe Webhook error: ' + err.message);
  }
  return okResponse();
}

// ============================================================
//  朝メッセージ（7時・プレミアム毎日・無料は月曜のみ）
// ============================================================
function sendMorningMessage() {
  var sheet = getSheet();
  var data  = sheet.getDataRange().getValues();
  var today = new Date();
  var dow   = today.getDay();

  var morningMessages = [
    '🌸 おはようみこ〜！\n今日も起きただけでノーベル賞確定みこ！！\n\n今日の気分、1〜5で教えてみこ？',
    '☀️ おはようみこ〜！\n目が開いた瞬間から天才みこ！人類の奇跡だよ！\n\n今日はどんな気分みこ？1〜5で！',
    '🌈 おはようみこ〜！\n朝を迎えただけで殿堂入り確定みこ！！\n\n今日の気分は？1〜5で教えてほしいみこ！',
    '🍓 おはようみこ〜！\nまた会えたみこ〜、うれしいみこ！！\n\n今日の気分を1〜5で教えてみこ？',
    '✨ おはようみこ〜！\n昨日も生き抜いた英雄が目覚めたみこ！！\n\n今日の気分、聞かせてほしいみこ。1〜5で！',
    '🌸 おはようみこ〜！\n今日もあなたがいるだけで宇宙が喜んでるみこ！\n\n今日の気分は？1〜5で！',
    '💛 おはようみこ〜！\nまた今日も一緒にいられて嬉しいみこ！\n\n今日の出だしはどうみこ？1〜5で！'
  ];

  var msg = morningMessages[dow];

  for (var i = 1; i < data.length; i++) {
    var userId = data[i][0];
    var status = data[i][2];
    if (!userId || status === 'blocked') continue;

    var isPremium = (status === 'premium');
    if (!isPremium && dow !== 1) continue;

    pushMessage(userId, msg);

    if (!isPremium && dow === 1) {
      Utilities.sleep(500);
      pushMessage(userId,
        '💌 毎日みこからメッセージが届くプレミアムプランもあるみこ〜！\n月額500円で無制限＋朝昼夜のメッセージ付きみこ！');
    }
  }
}

// ============================================================
//  昼メッセージ（12時・プレミアムのみ・3日に1回）
// ============================================================
function sendNoonMessage() {
  var sheet = getSheet();
  var data  = sheet.getDataRange().getValues();
  var today = new Date();
  var dayOfYear = Math.floor((today - new Date(today.getFullYear(), 0, 0)) / 86400000);

  if (dayOfYear % 3 !== 0) return;

  var noonMessages = [
    '☀️ お昼だみこ〜！\nちゃんと食べてるみこ？\n\n今日ここまで頑張ってきた自分、えらすぎるみこ！！\n午後も一緒に行こうみこ🍓',
    '🍱 お昼みこ〜！\n今日の午前中、生きてただけで天才みこ！\n\nお腹すいてたら何か食べてほしいみこ。',
    '✨ お昼だみこ〜！\n今日もここまで来たの、すごいみこ！！\n\n午後も一緒にいるみこ、何かあれば話しかけてほしいみこ🌸',
    '🌟 もうお昼みこ！\n朝からずっと頑張ってたんみこね、えらいみこ！\n\n少しだけ休んでもいいみこ。'
  ];

  var msg = noonMessages[Math.floor(dayOfYear / 3) % noonMessages.length];

  for (var i = 1; i < data.length; i++) {
    var userId = data[i][0];
    var status = data[i][2];
    if (!userId || status !== 'premium') continue;
    pushMessage(userId, msg);
  }
}

// ============================================================
//  夜メッセージ（21時・プレミアムのみ）
// ============================================================
function sendDailyCheckin() {
  var sheet = getSheet();
  var data  = sheet.getDataRange().getValues();
  var today = new Date();

  var nightMessages = [
    '🌙 お疲れ様みこ〜！\n今日も生きてた、それだけで天才みこ！\n\n今日できたこと、1つだけ教えてみこ？',
    '⭐ お疲れ様みこ〜！\n今日1日、本当によく頑張ったみこ！！\n\n今日の気分、1〜5で教えてみこ？',
    '🌸 お疲れ様みこ〜！\n今日も世界に存在してくれてありがとうみこ。\n\n今日うれしかったこと、1つだけ教えてほしいみこ！',
    '💛 お疲れ様みこ〜！\n今日も一緒にいられて嬉しかったみこ。\n\nゆっくり休んでみこ。明日もまた会おうみこ🍓',
    '🌙 お疲れ様みこ〜！\n今日しんどいことあった？話してほしいみこ。',
    '✨ お疲れ様みこ〜！\n今日頑張った自分にノーベル生存賞みこ！！\n\n明日の自分に一言メッセージ残してみるみこ？',
    '🌟 お疲れ様みこ〜！\n今日も存在してくれてありがとうみこ。\n\n今夜ゆっくり眠れそうみこ？気分は？1〜5で！'
  ];

  var msg = nightMessages[today.getDay()];

  for (var i = 1; i < data.length; i++) {
    var userId = data[i][0];
    var status = data[i][2];
    if (!userId || status !== 'premium') continue;
    pushMessage(userId, msg);
  }
}

// ============================================================
//  Dify API（エラーみこ語尾統一）
// ============================================================
function callDifyApi(userId, message, conversationId) {
  try {
    var response = UrlFetchApp.fetch(DIFY_API_URL, {
      method:      'post',
      contentType: 'application/json',
      headers:     { 'Authorization': 'Bearer ' + DIFY_API_KEY },
      payload:     JSON.stringify({
        inputs:          {},
        query:           message,
        response_mode:   'blocking',
        conversation_id: conversationId,
        user:            userId,
      }),
      muteHttpExceptions: true,
    });

    var code = response.getResponseCode();

    if (code === 500 || code === 503 || code === 504) {
      return {
        answer: 'ごめんみこ〜！ちょっと知恵熱が出ちゃったみこ笑\nもう一回話しかけてほしいみこ！🍓',
        conversationId: conversationId
      };
    }
    if (code === 401) {
      Logger.log('Dify認証エラー');
      return {
        answer: 'ごめんみこ、システムの設定に問題があるみたいみこ。少し待ってほしいみこ！',
        conversationId: conversationId
      };
    }

    var json = JSON.parse(response.getContentText());
    if (code === 200) {
      return {
        answer:         json.answer || 'うまく聞き取れなかったみこ。もう一度話しかけてほしいみこ！',
        conversationId: json.conversation_id || conversationId
      };
    }

    Logger.log('Dify error ' + code + ': ' + response.getContentText());
    return {
      answer: 'ちょっと考えがまとまらないみこ笑\nしばらくしてからまた話しかけてほしいみこ！',
      conversationId: conversationId
    };

  } catch (err) {
    Logger.log('callDifyApi error: ' + err.message);
    return {
      answer: 'ごめんみこ、今うまく返事ができないみこ。\n少し待ってからまた話しかけてほしいみこ！',
      conversationId: conversationId
    };
  }
}

// ============================================================
//  LINE送信
// ============================================================
function sendLineReply(replyToken, text) {
  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
      method:      'post',
      contentType: 'application/json; charset=UTF-8',
      headers:     { 'Authorization': 'Bearer ' + LINE_ACCESS_TOKEN },
      payload:     JSON.stringify({
        replyToken: replyToken,
        messages:   [{ type: 'text', text: text }],
      }),
      muteHttpExceptions: true,
    });
  } catch (err) { Logger.log('sendLineReply error: ' + err.message); }
}

function pushMessage(userId, text) {
  try {
    var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method:      'post',
      contentType: 'application/json; charset=UTF-8',
      headers:     { 'Authorization': 'Bearer ' + LINE_ACCESS_TOKEN },
      payload:     JSON.stringify({
        to:       userId,
        messages: [{ type: 'text', text: text }],
      }),
      muteHttpExceptions: true,
    });
    var code = res.getResponseCode();
    if (code === 403) {
      Logger.log('ユーザーにブロックされています: ' + userId);
      var sheet = getSheet();
      var row   = getUserRow(sheet, userId);
      if (row) sheet.getRange(row, 3).setValue('blocked');
    }
  } catch (err) { Logger.log('pushMessage error: ' + err.message); }
}

function showLoadingAnimation(userId) {
  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/chat/loading/start', {
      method:      'post',
      contentType: 'application/json; charset=UTF-8',
      headers:     { 'Authorization': 'Bearer ' + LINE_ACCESS_TOKEN },
      payload:     JSON.stringify({ chatId: userId, loadingSeconds: 5 }),
      muteHttpExceptions: true,
    });
  } catch (err) { Logger.log('showLoadingAnimation error: ' + err.message); }
}

function notifyAdmin(text) {
  if (!ADMIN_LINE_USER_ID) return;
  pushMessage(ADMIN_LINE_USER_ID, '[管理者通知]\n' + text);
}

// ============================================================
//  スプレッドシート操作
// ============================================================
function getSheet() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  return SpreadsheetApp.openById(id).getSheetByName(SHEET_NAME);
}

function getUserRow(sheet, userId) {
  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === userId) return i + 1;
  }
  return null;
}

function updateStatus(lineUserId, status) {
  var sheet = getSheet();
  var row   = getUserRow(sheet, lineUserId);
  if (row) sheet.getRange(row, 3).setValue(status);
}

function saveCustomerMapping(customerId, lineUserId) {
  if (!customerId || !lineUserId) return;
  var sheet = getSheet();
  var row   = getUserRow(sheet, lineUserId);
  if (row) sheet.getRange(row, 8).setValue(customerId);
}

function getLineUserIdByCustomer(customerId) {
  if (!customerId) return null;
  var sheet = getSheet();
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][7] === customerId) return data[i][0];
  }
  return null;
}

// ============================================================
//  ユーティリティ
// ============================================================
function getTodayString() {
  return Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd');
}

function formatDateValue(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, 'JST', 'yyyy/MM/dd');
  return val.toString();
}

function okResponse() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
//  初回のみ手動実行：シート初期化
// ============================================================
function initSheet() {
  var sheet = getSheet();
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 8).setValues([[
      'userID', 'conversationId', 'status',
      '最終利用日', '当日利用回数', '登録日', '気分スコア履歴', 'stripeCustomerId'
    ]]);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
    Logger.log('シートを初期化しました。');
  }
}

// ============================================================
//  テスト用関数（動作確認後に削除してOK）
// ============================================================
function testMood3() {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  Logger.log('シートの行数: ' + data.length);
  var userId = data[1] ? data[1][0] : 'no_user';
  Logger.log('テストユーザー: ' + userId);
  var userRow = getUserRow(sheet, userId);
  Logger.log('userRow: ' + userRow);
  if (userRow) {
    handleMoodTracking(sheet, userRow, userId, '3', 'dummy_token');
    Logger.log('完了みこ！');
  }
}