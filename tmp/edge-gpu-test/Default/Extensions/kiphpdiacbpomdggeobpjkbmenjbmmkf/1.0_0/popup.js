document.addEventListener("DOMContentLoaded", function () {
  const clicksInput = document.getElementById("clicksCount");
  const randomToggle = document.getElementById("randomToggle");
  const startButton = document.getElementById("startClicking");
  const prefixTextarea = document.getElementById("prefix");

  console.log(clicksInput);
  console.log(randomToggle);

  // 保存されたクリック数を読み込む
  chrome.storage.local.get(["clicksCount"], function (result) {
    if (result.clicksCount) {
      clicksInput.value = result.clicksCount;
    } else {
      clicksInput.value = 1;
      console.log("clicksInput.value" + clicksInput.value);
    }
  });

  // 保存されたランダムトグルを読み込む
  chrome.storage.local.get(["randomToggle"], function (result) {
    if (result.randomToggle !== undefined) {
      // 保存された状態があれば、チェックボックスに適用
      randomToggle.checked = result.randomToggle;
      prefixTextarea.disabled = false;
    } else {
      prefixTextarea.disabled = true;
    }
  });

  // 保存された接頭語を読み込む
  chrome.storage.local.get("prefix", function (result) {
    if (result.prefix !== undefined) {
      prefixTextarea.value = result.prefix;
    }
  });

  // 接頭語が変更されたときに保存
  prefixTextarea.addEventListener("change", function () {
    chrome.storage.local.set({ prefix: prefixTextarea.value });
  });

  // チェックボックスの状態が変更されたときに保存
  randomToggle.addEventListener("change", function () {
    chrome.storage.local.set({ randomToggle: randomToggle.checked });
    prefixTextarea.disabled = !this.checked;
  });

  startButton.addEventListener(
    "click",
    function () {
      console.log(clicksInput);
      console.log(randomToggle);
      console.log(clicksInput.value);
      console.log(randomToggle.checked);
      const clicksCount = clicksInput.value;
      const isRandomToggle = randomToggle.checked;
      chrome.storage.local.set({ clicksCount: clicksCount }, function () {
        console.log("クリック回数が保存されました: ", clicksCount);
      });
      chrome.storage.local.set({ randomToggle: isRandomToggle }, function () {
        console.log("ランダムトグルが保存されました: ", isRandomToggle);
      });

      chrome.tabs.query({ active: true, currentWindow: true }, function (tabs) {
        chrome.tabs
          .sendMessage(tabs[0].id, {
            tabId: tabs[0].id,
            action: "startClicking",
            clicksCount: clicksCount,
            randomToggle: isRandomToggle,
          })
          .catch((error) => {
            console.log(error);
          });
      });
    },
    false
  );
});
