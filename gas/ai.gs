/**
 * 既存 Code.gs の認証済みAPI振り分けから呼び出すAI処理。
 * 成功時は文字列を返す。失敗時はコードだけを持つ Error を throw する。
 * 呼び出し元は error.message を {ok:false,error:"コード"} に変換すること。
 * 成功時は既存APIと同様に {ok:true,data:返答テキスト} で返すこと。
 * APIキー、入力、レスポンス全文はログに出さない。
 */
function aiEnabled_() {
  return !!(PropertiesService.getScriptProperties()
    .getProperty('ANTHROPIC_API_KEY') || '').trim();
}

function aiAction_(payload) {
  try {
    var properties = PropertiesService.getScriptProperties();
    var key = (properties.getProperty('ANTHROPIC_API_KEY') || '').trim();
    if (!key) throw new Error('no_key');

    var prompt = payload && payload.prompt;
    if (typeof prompt !== 'string' || !prompt.trim()) throw new Error('ai_error');
    // JavaScriptの文字列長で上限を判定し、外部送信前に拒否する。
    if (prompt.length > 80000) throw new Error('too_long');

    var model = (properties.getProperty('AI_MODEL') || '').trim() || 'claude-sonnet-5-5';
    var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'x-api-key': key,
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify({
        model: model,
        max_tokens: 8192,
        messages: [{ role: 'user', content: prompt }]
      }),
      muteHttpExceptions: true
    });

    var status = response.getResponseCode();
    if (status === 429 || status === 529) throw new Error('rate_limited');
    if (status < 200 || status >= 300) throw new Error('ai_error');

    var data = JSON.parse(response.getContentText());
    if (!data || typeof data !== 'object') throw new Error('ai_error');
    if (data.stop_reason === 'refusal') throw new Error('refused');
    // 出力上限などで途中終了した結果を確認画面へ渡さない。
    if (data.stop_reason !== 'end_turn' || !Array.isArray(data.content)) {
      throw new Error('ai_error');
    }
    var text = data.content.filter(function (block) {
      return block && block.type === 'text' && typeof block.text === 'string';
    }).map(function (block) { return block.text; }).join('\n');
    if (!text.trim()) throw new Error('ai_error');
    return text;
  } catch (error) {
    // 通信例外やJSON解析エラーの詳細をクライアントに漏らさない。
    var code = error && error.message;
    var allowed = ['no_key', 'rate_limited', 'refused', 'too_long', 'ai_error'];
    throw new Error(allowed.indexOf(code) !== -1 ? code : 'ai_error');
  }
}
