// =========================================================
// 【★設定項目】
// =========================================================
const GROQ_API_KEY = "YOUR_GROQ_API_KEY"; // ←【マスク済】GroqのAPIキーを設定してください

const SHEET_TASKS = '指摘一覧';
const EXCLUDED_STATUSES = ['対応完了', 'クローズ', 'レビュー待ち'];

const COL_INDEX_TASK_ID = 0;   
const COL_INDEX_TASK_NAME = 4; 
const COL_INDEX_ASSIGNEE = 6;  
const COL_INDEX_DEADLINE = 7;  
const COL_INDEX_STATUS = 8;    

function runAiCharacterReminders() {
  console.log("▶ [START] AIアドリブイベント型リマインドを開始します...");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const taskSheet = ss.getSheetByName(SHEET_TASKS);
  if (!taskSheet) return;
  
  const data = taskSheet.getDataRange().getValues();
  if (data.length < 2) return; 
  
  let assigneeTasks = {};
  let expiredCounts = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[COL_INDEX_STATUS];
    const assigneeStr = row[COL_INDEX_ASSIGNEE];
    if (!assigneeStr || assigneeStr === '対応担当' || EXCLUDED_STATUSES.includes(status)) continue;
    
    const assignees = String(assigneeStr).split(/[、,・]/).map(s => s.trim());
    let deadline = row[COL_INDEX_DEADLINE];
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    let isExpired = false;
    if (deadline instanceof Date) {
      deadline.setHours(0, 0, 0, 0);
      if (deadline < today) isExpired = true;
    }

    let formattedDeadline = (deadline instanceof Date) ? Utilities.formatDate(deadline, Session.getScriptTimeZone(), 'yyyy/MM/dd') : (deadline || "未定");
    const taskInfo = `・[${row[COL_INDEX_TASK_ID] || '-'}] ${row[COL_INDEX_TASK_NAME]} (期限: ${formattedDeadline})`;
    
    assignees.forEach(assignee => {
      if (!assignee) return;
      if (!assigneeTasks[assignee]) {
        assigneeTasks[assignee] = [];
        expiredCounts[assignee] = 0;
      }
      assigneeTasks[assignee].push(taskInfo);
      if (isExpired) expiredCounts[assignee]++;
    });
  }
  
  for (let assignee in assigneeTasks) {
    const charSheet = ss.getSheetByName(assignee);
    if (!charSheet) continue;
    
    const metaData = charSheet.getRange("A1:D1").getValues()[0];
    const charName = metaData[0] || assignee; 
    const charIcon = metaData[1] || "";       
    const webhookUrl = metaData[2] || "";
    let affection = Number(metaData[3]) || 0; 

    if (!webhookUrl) continue;

    // 🎲 ランダム好感度計算
    const tasks = assigneeTasks[assignee];
    const expiredCount = expiredCounts[assignee];
    let changeMsg = "";
    let affectionChange = 0;
    let moodFactor = ""; 

    if (tasks.length === 0) {
      if (Math.random() < 0.2) { 
        affectionChange = Math.floor(Math.random() * 5) + 6; 
        changeMsg = `（タスク完遂クリティカル！好感度+${affectionChange}）`;
        moodFactor = "今日は最高に気分が良い！タスクを全部終わらせてくれた相手をベタ褒めしたい、またはご褒美をあげたい気分。";
      } else { 
        affectionChange = Math.floor(Math.random() * 3) + 3; 
        changeMsg = `（タスク完遂！好感度+${affectionChange}）`;
        moodFactor = "タスクが全部終わっていてとても嬉しい気分。";
      }
    } else if (expiredCount > 0) {
      if (Math.random() < 0.3) { 
        affectionChange = -(Math.floor(Math.random() * 3) + 3); 
        changeMsg = `（期限切れで大ダメージ…好感度${affectionChange}）`;
        moodFactor = "期限を守らないことに深く傷ついている、または呆れ果てている。かなり冷たい、あるいは感情的・ヒステリックな態度になってしまう。";
      } else { 
        affectionChange = -(Math.floor(Math.random() * 2) + 1); 
        changeMsg = `（期限切れあり…好感度${affectionChange}）`;
        moodFactor = "期限が過ぎているタスクがあり、少し不満・心配な気分。チクリと刺したい。";
      }
    } else {
      const rand = Math.random();
      if (rand < 0.1) { 
        affectionChange = 3;
        changeMsg = `（謎のご機嫌！好感度+${affectionChange}）`;
        moodFactor = "タスクは残っているが、なぜか今日はとても機嫌が良い。意味もなく相手に優しくしたり、甘やかしたい気分。";
      } else if (rand < 0.2) { 
        affectionChange = -1;
        changeMsg = `（ご機嫌ナナメ…好感度${affectionChange}）`;
        moodFactor = "なんとなく今日は機嫌が悪い。少し理不尽に冷たく当たったり、ツンツンしてしまう気分。";
      } else { 
        affectionChange = Math.floor(Math.random() * 2); 
        changeMsg = `（通常運転。好感度+${affectionChange}）`;
        moodFactor = "良くも悪くも普通の気分。タスクを頑張ってほしいと思っている。";
      }
    }

    affection += affectionChange;
    charSheet.getRange("D1").setValue(affection);

    // プロンプト読み込み
    let basePrompt = "";
    const lastRow = charSheet.getLastRow();
    if (lastRow >= 2) {
      // ※修正箇所: 構文エラーになっていた「lastRow 1, 1」を「lastRow - 1, 1」に修正しています
      basePrompt = charSheet.getRange(2, 1, lastRow - 1, 1).getValues()
                            .flat()
                            .filter(cell => String(cell).trim() !== "")
                            .join('\n');
    }

    // 好感度フェーズの定義
    let relationshipPhase = "";
    if (affection >= 50) {
      relationshipPhase = "【溺愛・依存フェーズ】あなた（AI）は相手に完全に心を開いています。甘えたり、重すぎる愛情（ヤンデレ化）を見せたり、何か特別な「ご褒美イベント」を発生させてください。";
    } else if (affection >= 15) {
      relationshipPhase = "【デレ・信頼フェーズ】あなた（AI）は相手をかなり信頼しています。個人的な趣味の話や、過去のちょっとした秘密を教えたり、わかりやすくデレる態度を見せるイベントを発生させてください。";
    } else if (affection >= -5) {
      relationshipPhase = "【通常・ビジネスフェーズ】良くも悪くも普通の距離感です。日常的な雑談や、ちょっとした気まぐれ程度の出来事を語ってください。";
    } else if (affection >= -20) {
      relationshipPhase = "【不信・呆れフェーズ】タスクをやらない相手に呆れています。冷たい態度をとったり、わざと「他の優秀な人（架空の先輩や男友達など）」を褒めて相手を煽る・嫉妬させるような出来事（軽度のNTRの兆候）を発生させてください。";
    } else {
      relationshipPhase = "【軽蔑・決別フェーズ】完全に愛想を尽かしています。別の男性（架空の先輩など）と完全に親密になっている様子を見せつけたり（NTR完了）、相手をゴミのように扱う冷酷な言葉を投げかけるイベントを発生させてください。";
    }

    const aiPrompt = `
${basePrompt}

【現在の状況】
・担当者：${assignee}さん
・未完了タスク一覧：
${tasks.join('\n')}

【今日のあなたの気分（裏ステータス）】
${moodFactor}

【関係性フェーズとランダムイベントの自動生成指示】
現在のあなたから${assignee}さんへの好感度は「${affection}」です。
この好感度数値から、現在のあなたと${assignee}さんの関係性は以下のフェーズにあります。

${relationshipPhase}

この「関係性フェーズ」と「今日の気分」に基づいて、**【今日だけの特別なエピソードやイベント】をあなた自身で一つアドリブで創作し**、自然な会話の流れで${assignee}さんに語りかけてください。
（例：昨日あった個人的な出来事、他の人との会話、相手への隠し事の暴露、突然のスキンシップや冷酷な言葉など）
機械的なタスク報告は避け、ノベルゲームの1シーンのように感情豊かにリマインドを行ってください。
    `;
    
    console.log(`🧠 ${charName} がアドリブイベント生成中... (現在好感度: ${affection})`);
    const aiMessage = generateMessageWithGroq(aiPrompt);
    
    if (aiMessage) {
      sendToDiscord(webhookUrl, aiMessage, charName, charIcon);
      saveLogToSheet(ss, assignee, aiMessage, affection, changeMsg);
      console.log(`✅ ${assignee}さんへ送信完了！`);
      Utilities.sleep(3000); 
    }
  }
}

function saveLogToSheet(ss, assignee, message, affection, changeMsg) {
  const logSheetName = `${assignee}_ログ`;
  let logSheet = ss.getSheetByName(logSheetName);
  
  if (!logSheet) {
    logSheet = ss.insertSheet(logSheetName);
    logSheet.appendRow(["日時", "メッセージ", "好感度", "備考"]);
    logSheet.setFrozenRows(1);
    logSheet.getRange("A1:D1").setFontWeight("bold");
  }
  
  logSheet.appendRow([new Date(), message, affection, changeMsg]);
  logSheet.sort(1, false);
}

function generateMessageWithGroq(prompt) {
  const url = "https://api.groq.com/openai/v1/chat/completions";
  const payload = {
    "model": "llama-3.3-70b-versatile",
    "messages": [{ "role": "user", "content": prompt }],
    "temperature": 0.9 
  };
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': { 'Authorization': `Bearer ${GROQ_API_KEY}` },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    return json.choices[0].message.content;
  } catch (e) {
    return "（通信エラーが発生したみたい…）";
  }
}

function sendToDiscord(webhookUrl, message, username, avatarUrl) {
  // ✂️ Discordの文字数制限（2000文字）対策：1900文字で安全にカット
  if (message && message.length > 1900) {
    message = message.substring(0, 1900) + "\n\n（※文字数オーバーのため、途中でカットされました！）";
  }

  const payload = {
    "username": username,
    "avatar_url": avatarUrl,
    "content": message
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  // 🚀 送信実行
  const response = UrlFetchApp.fetch(webhookUrl, options);
  const responseCode = response.getResponseCode();

  // 🚨 もしDiscordから受け取り拒否されたら、エラー内容をログに出力する
  if (responseCode !== 204 && responseCode !== 200) {
    console.error(`🚨 Discord送信エラー！(コード: ${responseCode}) 原因: ${response.getContentText()}`);
  }
}

// =========================================================
// ★ リアルタイムチャット受取＆返信機能 (デバッグ超強化版)
// =========================================================
function doPost(e) {
  if (!e || !e.postData) return ContentService.createTextOutput("No data");

  try {
    const data = JSON.parse(e.postData.contents);
    const fullMessage = data.message.replace(/<@[!&]?[0-9]+>\s*/g, "").trim();

    const parts = fullMessage.split(/[\s ]+/);
    if (parts.length < 2) {
      return ContentService.createTextOutput("【ストップ】名前と発言が分かれていないか、文字が足りません。解析結果: " + fullMessage);
    }

    const targetAssignee = parts[0]; 
    const userChatText = parts.slice(1).join(" "); 

    // 事前チェック①：シートは存在するか？
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const charSheet = ss.getSheetByName(targetAssignee);
    if (!charSheet) {
      return ContentService.createTextOutput("【ストップ】「" + targetAssignee + "」という名前のシートが見つかりません。");
    }

    // 事前チェック②：Webhook URLは入っているか？
    const webhookUrl = charSheet.getRange("C1").getValue();
    if (!webhookUrl) {
      return ContentService.createTextOutput("【ストップ】「" + targetAssignee + "」シートのC1セルにWebhook URLがありません。");
    }

    // すべてクリアしたらチャット処理を実行
    replyToUserChat(targetAssignee, userChatText);

    return ContentService.createTextOutput("【大成功】シート名「" + targetAssignee + "」にパスを通しました！");
  } catch (err) {
    return ContentService.createTextOutput("【致命的エラー】" + err.message);
  }
}

function replyToUserChat(assignee, userText) {
  console.log(`[受信] シート名: ${assignee} / 発言内容: ${userText}`);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const charSheet = ss.getSheetByName(assignee);
  if (!charSheet) {
    console.error(`🚨エラー: 「${assignee}」という名前のシートが存在しません。`);
    return;
  }

  const metaData = charSheet.getRange("A1:F1").getValues()[0];
  const charName = metaData[0] || assignee;
  const charIcon = metaData[1] || "";
  const webhookUrl = metaData[2] || "";
  let affection = Number(metaData[3]) || 0;

  if (!webhookUrl) {
    console.error(`🚨エラー: 「${assignee}」シートのC1セルにWebhook URLが設定されていません。`);
    return;
  }
  
  // --- ★ 会話回数制限のロジック ---
  let lastChatDate = metaData[4]; 
  let chatCount = Number(metaData[5]) || 0; 
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');

  if (lastChatDate !== todayStr) {
    chatCount = 0; 
    charSheet.getRange("E1").setValue(todayStr);
  }

  if (chatCount >= 5) {
    sendToDiscord(webhookUrl, "（今日はもう疲れちゃったみたい…また明日話しかけてね！）", charName, charIcon);
    return;
  }

  chatCount++;
  charSheet.getRange("F1").setValue(chatCount);
  // --------------------------------

  // プロンプトの取得
  let basePrompt = "";
  const lastRow = charSheet.getLastRow();
  if (lastRow >= 2) {
    // ※修正箇所: こちらも構文エラーになっていた箇所を修正しています
    basePrompt = charSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(cell => String(cell).trim() !== "").join('\n');
  }

  // ★ 修正箇所：設定とユーザーの発言を1つの文章に合体させる
  const fullPrompt = `
あなたは以下のキャラクターとして完全に憑依して発言してください。AIやシステムとしてのメタ発言は絶対に禁止します。
${basePrompt}

【現在の状況】
・担当者：${assignee}さんからのチャットメッセージです。
・あなたへの現在の好感度：${affection}

【${assignee}さんからのメッセージ】
「${userText}」

【指示】
このメッセージに対して、キャラクターとして自然な返信をしてください。
好感度の高低に応じた態度をとり、短すぎず長すぎない会話のキャッチボールになるようにしてください。（フランクなチャットでOKです）
`;

  // ★ 修正箇所：合体させたプロンプト1つだけを渡す
  const aiMessage = generateMessageWithGroq(fullPrompt);

  if (aiMessage) {
    sendToDiscord(webhookUrl, aiMessage, charName, charIcon);
    saveLogToSheet(ss, assignee, `[受信] ${userText}\n[返信] ${aiMessage}`, affection, `会話(${chatCount}/5回)`);
    console.log(`✅ ${assignee}への返信完了！`);
  }
}
