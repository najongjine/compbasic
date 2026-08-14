chrome.runtime.onMessage.addListener(function (request, sender, sendResponse) {
  console.log("background test");
  if (request.allClicksDone) {
    console.log("All clicks have been completed.");
    // 例: 簡単な通知を送る
    chrome.notifications.create(
      "",
      {
        // 空文字列で自動ID生成
        type: "basic",
        iconUrl: "images/icon128.png", // 適切なアイコンURLに変更してください
        title: "クリック完了",
        message: "指定されたすべてのクリックが完了しました！",
      },
      function (notificationId) {
        console.log("Notification created with ID:", notificationId);
        console.log("Tab ID:", request.tabId);
        chrome.storage.local.set({ [notificationId]: request.tabId });
      }
    );
  }
});

chrome.notifications.onClicked.addListener(function (notificationId) {
  // 保存されていたタブIDを取得
  chrome.storage.local.get([notificationId], function (result) {
    if (result[notificationId]) {
      var tabId = result[notificationId];
      // タブをアクティブにする
      chrome.tabs.update(tabId, { active: true });
      // オプションでウィンドウもフォーカス
      chrome.tabs.get(tabId, function (tab) {
        chrome.windows.update(tab.windowId, { focused: true });
      });
      // 用済みの通知IDをクリア
      chrome.storage.local.remove([notificationId]);
    }
  });
});
