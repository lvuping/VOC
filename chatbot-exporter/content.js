function extractChatContent() {
  // Streamlit 채팅 메시지를 선택
  const messages = document.querySelectorAll('.stChatMessage');
  let markdown = '';
  
  messages.forEach((message) => {
    // 사용자와 AI 메시지를 구분
    const isUser = message.querySelector('.stChatMessageContent.user');
    const content = message.querySelector('.stMarkdown p').textContent;
    
    if (isUser) {
      markdown += `\n### Human:\n${content}\n`;
    } else {
      markdown += `\n### Assistant:\n${content}\n`;
    }
  });
  
  return markdown;
}

// popup.js와 통신하기 위한 리스너
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'exportChat') {
    const chatContent = extractChatContent();
    sendResponse({ content: chatContent });
  }
}); 