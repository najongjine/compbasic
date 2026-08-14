function clickButtonOnceWhenActivated(targetButton, callback) {
  // ボタンの変更を監視するためのオブザーバーを作成
  const observer = new MutationObserver((mutations) => {
    mutations.forEach((mutation) => {
      if (
        mutation.type === "attributes" &&
        mutation.attributeName === "disabled" &&
        !targetButton.disabled
      ) {
        callback(targetButton);
        observer.disconnect();
      }
    });
  });

  // オブザーバーの設定
  const config = { attributes: true, attributeFilter: ["disabled"] };

  // ボタンの監視を開始
  observer.observe(targetButton, config);
}

// クリックする回数と現在のカウントを管理
let targetClicks = 0;
let currentClicks = 0;

async function tryClicking(clicksCount, tabId, randomToggle) {
  targetClicks = clicksCount;
  currentClicks = 0;

  const targetButton = getButtonFromText("１枚のみ生成");
  if (randomToggle) {
    // 特定の2つのボタンをクリックする
    await clickSpecialButtons();
    // その後、通常のクリックプロセスを続ける
    targetButton.click();
    currentClicks++;
  } else {
    // トグルがOFFの場合は、通常のクリックプロセスだけを行う
    targetButton.click();
    currentClicks++;
  }

  clickButtonOnce(targetButton, tabId, randomToggle);
}

function clickButtonOnce(targetButton, tabId, randomToggle) {
  clickButtonOnceWhenActivated(targetButton, async (button) => {
    // 指定された回数に達するまで、再活性化を監視
    if (currentClicks < targetClicks) {
      if (randomToggle) {
        await clickSpecialButtons();
      }
      setTimeout(() => {
        button.click();
        currentClicks++;
        console.log(`Clicked ${currentClicks} times`);
        clickButtonOnce(targetButton, tabId, randomToggle);
      }, 10000);
    } else {
      console.log("All clicks done.");
      // 全てのクリックが完了したことをbackground.jsに通知
      chrome.runtime.sendMessage({ allClicksDone: true, tabId: tabId });
    }
  });
}

function getButtonFromText(text) {
  var button = null;

  // ページ内のすべてのボタン要素を取得
  const buttons = document.querySelectorAll("button");

  // 各ボタンをチェックし、条件に合致するものを見つける
  buttons.forEach((b) => {
    // ボタン要素の子要素であるspanタグを探す
    const span = b.querySelector("span");
    // spanが存在し、そのテキストが「生成」であるかチェック
    if (span && span.textContent === text) {
      button = b;
    }
  });

  return button;
}

function clickSpecialButtons() {
  return new Promise((resolve, reject) => {
    const textarea = getTextArea();
    if (textarea) {
      textarea.value = "";
      const event = new Event("input", {
        bubbles: true,
        cancelable: true,
      });
      textarea.dispatchEvent(event);
    }

    setTimeout(() => {
      const randomButton = getButtonFromText("ランダム化");
      if (randomButton) {
        randomButton.click();
        setTimeout(() => {
          chrome.storage.local.get("prefix", function (result) {
            const newTextArea = getTextArea();
            let prefix = result.prefix || ""; // 保存された接頭語を取得、なければ空文字
            newTextArea.value = prefix + newTextArea.value; // 接頭語を追加
            dispatchEvents(newTextArea);
            console.log(newTextArea.value);
            resolve();
          });
        }, 500);
      } else {
        resolve();
      }
    }, 500);
  });
}

function dispatchEvents(element) {
  ["input", "change", "blur"].forEach((type) => {
    const event = new Event(type, { bubbles: true, cancelable: true });
    element.dispatchEvent(event);
  });
}

function getTextArea() {
  const specificString = "プロンプトを入力し、理想の画像を生成しましょう";
  var firstMatchedTextarea = null;

  // ページ内のすべてのtextarea要素を取得
  const textareas = document.querySelectorAll("textarea");

  // 条件に一致するtextareaを探す
  const matchedTextareas = Array.from(textareas).filter((textarea) =>
    textarea.placeholder.includes(specificString)
  );

  // 結果を確認（最初の一致したtextareaを例として使用）
  if (matchedTextareas.length > 0) {
    firstMatchedTextarea = matchedTextareas[0];
    console.log(firstMatchedTextarea); // または他の操作
  } else {
    console.log("一致するtextareaが見つかりませんでした。");
  }

  return firstMatchedTextarea;
}

// popup.jsからのメッセージを受け取る
chrome.runtime.onMessage.addListener(function (request, sender, sendResponse) {
  console.log("start clicking");
  if (request.action == "startClicking") {
    console.log("request.clicksCount" + request.clicksCount);
    console.log("request.tabId" + request.tabId);
    tryClicking(request.clicksCount, request.tabId, request.randomToggle);
  }
});
