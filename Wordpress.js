function postArticlesToWordPress() {
  try {
    const sheetName = '記事';
    const logSheetName = 'ログ';
    const keywordSheetName = 'キーワード';
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);
    const logSheet = spreadsheet.getSheetByName(logSheetName) || spreadsheet.insertSheet(logSheetName);
    const keywordSheet = spreadsheet.getSheetByName(keywordSheetName);

    if (!keywordSheet) {
      throw new Error(`シート「${keywordSheetName}」が見つかりません。`);
    }

    const wordpressUrl = keywordSheet.getRange('K1').getValue();
    const username = keywordSheet.getRange('K2').getValue();
    const password = keywordSheet.getRange('K3').getValue();
    const postStatus = keywordSheet.getRange('K4').getValue();
    const categoryData = keywordSheet.getRange('G2:H' + keywordSheet.getLastRow()).getValues()
      .filter(([id, category]) => id.toString().trim() !== '' && category.trim() !== '')
      .map(([id, category]) => ({ id: id.toString().trim(), name: category.trim() }));

    if (!wordpressUrl || !username || !password) {
      throw new Error('WordPress の設定が正しく入力されていません。キーワードシートを確認してください。');
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const articles = data.slice(1);

    if (logSheet.getLastRow() === 0) {
      logSheet.appendRow(['タイムスタンプ', 'モデル', 'プロンプト', 'レスポンス', 'トークン数']);
    }

    let postedCount = 0;

    articles.forEach((article, index) => {
      try {
        const seoTitle = article[headers.indexOf('SEOタイトル')];
        const slug = article[headers.indexOf('スラッグ')];
        const metaDescription = article[headers.indexOf('メタディスクリプション')];
        const metaKeyword = article[headers.indexOf('メタキーワード')];
        const content = article[headers.indexOf('記事')];
        const isPosted = article[headers.indexOf('WP投稿済み')];

        if (!seoTitle || !content) {
          console.log(`記事「${seoTitle || '(タイトルなし)'}」は必要な情報が不足しているためスキップします。`);
          return;
        }

        if (isPosted) {
          console.log(`記事「${seoTitle}」は既に投稿済みのためスキップします。`);
          return;
        }

        const status = postStatus.toLowerCase() === '公開' ? 'publish' : 'draft';

        // カテゴリーの選択
        const categoryNames = categoryData.map(category => category.name);
        const categoryPrompt = `以下の記事の内容に最も適したカテゴリーを、次のカテゴリーの中から最大3つ選んでください。複数選択する場合は、カンマまたは、で区切って記載して、余計なものは出力せずに回答だけ出力してください。\n\nカテゴリー:\n${categoryNames.join('\n')}\n\n記事内容:\n${content}\n\n最も適したカテゴリー:`;
        const messages = [
          { role: 'system', content: 'あなたは優秀なアシスタントです。' },
          { role: 'user', content: categoryPrompt }
        ];
        const categoryResponse = callAnthropicApi(messages, 'claude-3-haiku-20240307', logSheet);
        const selectedCategories = categoryResponse.trim().split(/,|、/).map(category => category.trim());

        // カテゴリーIDの取得
        const selectedCategoryIds = selectedCategories.map(selectedCategory => {
          const category = categoryData.find(category => category.name === selectedCategory);
          return category ? category.id : null;
        }).filter(id => id !== null);

        const postData = {
          title: seoTitle,
          slug: slug,
          content: content,
          status: status,
          categories: selectedCategoryIds,
          meta: {
            _yoast_wpseo_title: seoTitle,
            _yoast_wpseo_metadesc: metaDescription,
            _yoast_wpseo_focuskw: metaKeyword
          }
        };

        const requestOptions = {
          method: 'post',
          headers: {
            'Content-Type': 'application/json',
            'Authorization': 'Basic ' + Utilities.base64Encode(username + ':' + password)
          },
          payload: JSON.stringify(postData),
          muteHttpExceptions: true
        };

        const response = UrlFetchApp.fetch(wordpressUrl + '/wp-json/wp/v2/posts', requestOptions);
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();
        const timestamp = new Date().toLocaleString();

        if (responseCode === 201) {
          postedCount++;
          const responseJson = JSON.parse(responseText);
          const tokens = JSON.stringify(responseJson).length;
          logSheet.appendRow([timestamp, 'WordPress REST API', JSON.stringify(postData), JSON.stringify(responseJson), tokens]);
          sheet.getRange(index + 2, headers.indexOf('WP投稿済み') + 1).setValue(true);
          console.log(`記事「${seoTitle}」の投稿に成功しました。カテゴリー: ${selectedCategories.join(', ')}`);
        } else {
          const tokens = responseText.length;
          logSheet.appendRow([timestamp, 'WordPress REST API', JSON.stringify(postData), 'エラー: ' + responseText, tokens]);
          console.error(`記事「${seoTitle}」の投稿に失敗しました。レスポンスコード: ${responseCode}`);
        }
      } catch (error) {
        const timestamp = new Date().toLocaleString();
        const tokens = error.message.length;
        logSheet.appendRow([timestamp, 'WordPress REST API', '', 'リクエストエラー: ' + error.message, tokens]);
        console.error('Error posting article:', error);
      }
    });

    console.log(`記事の投稿が完了しました。投稿された記事の件数: ${postedCount}件`);
    SpreadsheetApp.getActiveSpreadsheet().toast(`記事の投稿が完了しました。投稿された記事の件数: ${postedCount}件`, 'WP記事投稿');
  } catch (error) {
    const timestamp = new Date().toLocaleString();
    const logSheetName = 'ログ';
    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(logSheetName) || SpreadsheetApp.getActiveSpreadsheet().insertSheet(logSheetName);
    const tokens = error.message.length;
    logSheet.appendRow([timestamp, 'スクリプトエラー', '', 'エラー: ' + error.message, tokens]);
    console.error('Error in script:', error);
    SpreadsheetApp.getActiveSpreadsheet().toast(`スクリプトエラー: ${error.message}`, 'エラー');
  }
}