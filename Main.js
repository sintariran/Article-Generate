const MAX_TOKEN_NUM = 2000;

function generateArticlesWithCombinedKeywords() {
  const keywordSheetName = 'キーワード';
  const articleSheetName = '記事';
  const logSheetName = 'ログ';

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const keywordSheet = spreadsheet.getSheetByName(keywordSheetName);
  const articleSheet = spreadsheet.getSheetByName(articleSheetName);
  const logSheet = spreadsheet.getSheetByName(logSheetName) || spreadsheet.insertSheet(logSheetName);

  if (!keywordSheet) {
    throw new Error(`シート「${keywordSheetName}」が見つかりません。`);
  }
  if (!articleSheet) {
    throw new Error(`シート「${articleSheetName}」が見つかりません。`);
  }

  const keyword1List = keywordSheet.getRange(2, 1, keywordSheet.getLastRow() - 1).getValues().flat().filter(String);
  const keyword2List = keywordSheet.getRange(2, 2, keywordSheet.getLastRow() - 1).getValues().flat().filter(String);

  let existingCombinations = [];
  if (articleSheet.getLastRow() > 1) {
    existingCombinations = articleSheet.getRange(2, 1, articleSheet.getLastRow() - 1, 2).getValues();
  }

  const articleLength = keywordSheet.getRange('E1').getValue() || 1000;
  const numSections = keywordSheet.getRange('E2').getValue() || 1;
  const maxArticlesNum = keywordSheet.getRange('E3').getValue() || 10;
  const modelName = keywordSheet.getRange('E4').getValue() || 'claude-3-haiku-20240307';
  const persona = keywordSheet.getRange('E5').getValue() || '';
  const customerJourney = keywordSheet.getRange('E6').getValue() || '';

  const systemPrompt = `あなたはSEOの記事生成のプロです。次のペルソナがカスタマージャーニーの情報収集や検討段階にネットから見つけて読みたくなる記事を作成して、HTML形式で出力してください。記事は${numSections}セクションに分けてください:\n\nペルソナ: ${persona}\nカスタマージャーニー: ${customerJourney}\n\n${keywordSheet.getRange('E7').getValue() || ''}`;

  const keywordCombinations = [];
  keyword1List.forEach(keyword1 => {
    keyword2List.forEach(keyword2 => {
      if (!existingCombinations.some(combination => combination[0] === keyword1 && combination[1] === keyword2)) {
        keywordCombinations.push([keyword1, keyword2]);
      }
    });
  });

  if (articleSheet.getLastRow() === 1) {
    articleSheet.getRange(1, 1, 1, 8).setValues([['キーワード1', 'キーワード2', 'SEOタイトル', 'スラッグ', 'メタディスクリプション', 'メタキーワード', '記事', 'チェック']]);
    articleSheet.getRange('G1').insertCheckboxes();    
  }

  if (logSheet.getLastRow() === 0) {
    logSheet.getRange(1, 1, 1, 5).setValues([['タイムスタンプ', 'モデル', 'プロンプト', 'レスポンス', 'トークン数']]);
  }

  const totalArticles = Math.min(keywordCombinations.length, maxArticlesNum);
  const showProgressToast = (elapsedSeconds) => {
    const progressMessage = `記事を生成しています... (${index + 1}/${totalArticles}) - ${elapsedSeconds}秒経過`;
    SpreadsheetApp.getActiveSpreadsheet().toast(progressMessage, "記事生成の進捗", 1);
  };  
  keywordCombinations.slice(0, totalArticles).forEach((combination, index) => {
    const [keyword1, keyword2] = combination;
    const prompt = `以下のキーワードを使って、${articleLength}文字程度、${numSections}セクションからなる記事を<section>タグを用いて作成してください:\n\nキーワード1: ${keyword1}\nキーワード2: ${keyword2}`;

    const messages = [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: prompt }
    ];

    const progressMessage = `記事を生成しています... (${index + 1}/${totalArticles})`;
    SpreadsheetApp.getActiveSpreadsheet().toast(progressMessage, "記事生成の進捗", 3);

    let article;
    if (modelName === 'gpt-4o') {
      article = callOpenAiApi(messages, modelName, logSheet, showProgressToast);
    } else if (modelName.startsWith('deepseek')) {
      article = callDeepSeekApi(messages, modelName, logSheet, showProgressToast);
    } else {
      article = callAnthropicApi(messages, modelName, logSheet, showProgressToast);
    }
    // SEOタイトル、メタディスクリプション、メタキーワードを生成
    const seoTitlePrompt = `以下の記事に適した1つの記事タイトルをSEOを踏まえて50文字以内で提案してください:\n\n${article}`;
    const seoTitle = modelName === 'gpt-4o' ? callOpenAiApi([{ role: 'user', content: seoTitlePrompt }], modelName, logSheet) : (modelName.startsWith('deepseek') ? callDeepSeekApi([{ role: 'user', content: seoTitlePrompt }], modelName, logSheet) : callAnthropicApi([{ role: 'user', content: seoTitlePrompt }], modelName, logSheet));
    const metaDescPrompt = `以下の記事に適したメタディスクリプションを120文字以内で提案してください:\n\n${article}`;
    const metaDesc = modelName === 'gpt-4o' ? callOpenAiApi([{ role: 'user', content: metaDescPrompt }], modelName, logSheet) : (modelName.startsWith('deepseek') ? callDeepSeekApi([{ role: 'user', content: metaDescPrompt }], modelName, logSheet) : callAnthropicApi([{ role: 'user', content: metaDescPrompt }], modelName, logSheet));
    const metaKeywordPrompt = `以下の記事に適したメタキーワードを10個以内、カンマ区切りで提案してください:\n\n${article}`;
    const metaKeyword = modelName === 'gpt-4o' ? callOpenAiApi([{ role: 'user', content: metaKeywordPrompt }], modelName, logSheet) : (modelName.startsWith('deepseek') ? callDeepSeekApi([{ role: 'user', content: metaKeywordPrompt }], modelName, logSheet) : callAnthropicApi([{ role: 'user', content: metaKeywordPrompt }], modelName, logSheet));
    // スラッグを生成
    const slugPrompt = `以下のSEOタイトルから適切なスラッグ(URLに使うスラッグ)を提案してください。結果だけを出力して:\n\n${seoTitle}`;
    const slug = modelName === 'gpt-4o' ? callOpenAiApi([{ role: 'user', content: slugPrompt }], modelName, logSheet) : (modelName.startsWith('deepseek') ? callDeepSeekApi([{ role: 'user', content: slugPrompt }], modelName, logSheet) : callAnthropicApi([{ role: 'user', content: slugPrompt }], modelName, logSheet));

    // 記事データをシートに追加  
    articleSheet.appendRow([keyword1, keyword2, seoTitle, slug, metaDesc, metaKeyword, article, '']);
    articleSheet.getRange(articleSheet.getLastRow(), 8).insertCheckboxes();    
  });
  SpreadsheetApp.getActiveSpreadsheet().toast("記事の生成が完了しました。", "完了", 3);
}