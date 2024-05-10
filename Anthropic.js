function callAnthropicApi(messages, modelName, logSheet, showProgressToast) {
  const systemMessage = messages.find(message => message.role === 'system');
  const userMessages = messages.filter(message => message.role === 'user');
  const assistantMessages = messages.filter(message => message.role === 'assistant');
  
  const messagesPayload = messages.filter(message => message.role !== 'system').map(message => ({
    "role": message.role,
    "content": message.content
  }));
  
  const scriptProperties = PropertiesService.getScriptProperties();
  const apiKey = scriptProperties.getProperty('ANTHROPIC_API_KEY');
  
  const requestOptions = {
    "method": "post",
    "headers": {
      "Content-Type": "application/json",
      "X-API-Key": apiKey,
      "anthropic-version": "2023-06-01"
    },
    "payload": JSON.stringify({
      "model": modelName,
      "max_tokens": MAX_TOKEN_NUM,
      "temperature": 0,
      ...(systemMessage && { "system": systemMessage.content }),
      "messages": messagesPayload
    }),
  };
  
  const maxRetries = 3;
  let retryCount = 0;
  
  const startTime = new Date().getTime();
  let elapsedSeconds = 0;
  
  while (retryCount < maxRetries) {
    try {
      const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", requestOptions);
      
      const json = JSON.parse(response.getContentText());
      if (json['type'] === 'message') {
        if (json['content'] && json['content'].length > 0 && json['content'][0]['text']) {
          const botReply = json['content'][0]['text'].trim();
          
          // APIのログをログシートに出力
          const timestamp = new Date().toLocaleString();
          const prompt = userMessages[0].content;
          const tokens = json['content'][0]['tokens'];
          logSheet.appendRow([timestamp, modelName, prompt, botReply, tokens]);
          
          return botReply;
        } else {
          console.error('Unexpected message format:', json);
          return 'すみません、今は適切な応答ができません。';
        }
      } else if (json['type'] === 'error') {
        console.error('Anthropic API error:', json['error']);
        if (json['error']['type'] === 'rate_limit_error') {
          retryCount++;
          if (retryCount < maxRetries) {
            const retryDelay = Math.pow(2, retryCount) * 1000;
            const waitingMessage = `APIリクエストがレート制限に達しました。${retryDelay}ミリ秒後に再試行します...`;
            SpreadsheetApp.getActiveSpreadsheet().toast(waitingMessage, "APIリクエスト待機中", 3);
            console.log(waitingMessage);
            Utilities.sleep(retryDelay);
          } else {
            throw new Error(`Anthropic API error: ${json['error']['message']}`);
          }
        } else {
          throw new Error(`Anthropic API error: ${json['error']['message']}`);
        }
      } else {
        console.error('Unexpected API response type:', json);
        throw new Error('Anthropic API returned an unexpected response type');
      }
      
      elapsedSeconds = Math.floor((new Date().getTime() - startTime) / 1000);
      showProgressToast(elapsedSeconds);
      Utilities.sleep(1000);
      
    } catch (error) {
      console.error('Error in callAnthropicApi:', error);
      throw error;
    }
  }
}

function callDeepSeekApi(messages, modelName, logSheet, showProgressToast) {
  const systemMessage = messages.find(message => message.role === 'system');
  const userMessages = messages.filter(message => message.role === 'user');
  const assistantMessages = messages.filter(message => message.role === 'assistant');
  
  const messagesPayload = messages.map(message => ({
    "role": message.role,
    "content": message.content
  }));
  
  const scriptProperties = PropertiesService.getScriptProperties();
  const apiKey = scriptProperties.getProperty('DEEPSEEK_API_KEY');
  
  const requestOptions = {
    "method": "post",
    "headers": {
      "Content-Type": "application/json",
      "Authorization": `Bearer ${apiKey}`
    },
    "payload": JSON.stringify({
      "model": modelName,
      "messages": messagesPayload
    }),
  };
  
  const maxRetries = 3;
  let retryCount = 0;
  
  const startTime = new Date().getTime();
  let elapsedSeconds = 0;
  
  while (retryCount < maxRetries) {
    try {
      const response = UrlFetchApp.fetch("https://api.deepseek.com/chat/completions", requestOptions);
      
      const json = JSON.parse(response.getContentText());
      if (json['choices'] && json['choices'].length > 0) {
        const botReply = json['choices'][0]['message']['content'].trim();
        
        // APIのログをログシートに出力
        const timestamp = new Date().toLocaleString();
        const prompt = userMessages[0].content;
        const tokens = json['usage']['total_tokens'];
        logSheet.appendRow([timestamp, modelName, prompt, botReply, tokens]);
        
        return botReply;
      } else {
        console.error('Unexpected message format:', json);
        return 'すみません、今は適切な応答ができません。';
      }
    } catch (error) {
      if (error.message.includes('Rate limit exceeded')) {
        retryCount++;
        if (retryCount < maxRetries) {
          const retryDelay = Math.pow(2, retryCount) * 1000;
          const waitingMessage = `APIリクエストがレート制限に達しました。${retryDelay}ミリ秒後に再試行します...`;
          SpreadsheetApp.getActiveSpreadsheet().toast(waitingMessage, "APIリクエスト待機中", 3);
          console.log(waitingMessage);
          Utilities.sleep(retryDelay);
        } else {
          console.error('Error in callDeepSeekApi:', error);
          throw error;
        }
      } else {
        console.error('Error in callDeepSeekApi:', error);
        throw error;
      }
    }
    
    elapsedSeconds = Math.floor((new Date().getTime() - startTime) / 1000);
    showProgressToast(elapsedSeconds);
    Utilities.sleep(1000);
  }
}