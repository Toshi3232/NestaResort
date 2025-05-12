function setOpenAIApiKey() {
  var apiKey = Browser.inputBox('OpenAI APIキーの設定', 'APIキーを入力してください:', Browser.Buttons.OK_CANCEL);
  if (apiKey !== 'cancel' && apiKey !== '') {
    PropertiesService.getScriptProperties().setProperty('OPENAI_API_KEY', apiKey);
    Browser.msgBox('APIキーが正常に設定されました。');
  } else {
    Browser.msgBox('APIキーの設定がキャンセルされました。');
  }
}