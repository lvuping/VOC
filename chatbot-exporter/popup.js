document.getElementById('exportBtn').addEventListener('click', async () => {
  // 현재 활성화된 탭에서 content script 실행
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  
  chrome.tabs.sendMessage(tab.id, { action: 'exportChat' }, (response) => {
    if (response && response.content) {
      // Blob 생성 및 다운로드
      const blob = new Blob([response.content], { type: 'text/markdown' });
      const url = URL.createObjectURL(blob);
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
      
      chrome.downloads.download({
        url: url,
        filename: `chat-export-${timestamp}.md`
      });
    }
  });
}); 