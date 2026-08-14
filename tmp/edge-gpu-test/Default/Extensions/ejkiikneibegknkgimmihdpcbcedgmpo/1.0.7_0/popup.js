var port = chrome.runtime.connect();
port.postMessage({action: 'start'});

var slide = document.getElementById('slide');
var button = document.getElementById('button');
const DEFAULT_VOLUME = 2;

// Retrieve the saved volume value when popup opens
chrome.storage.local.get('volumeLevel', function(data) {
  if (data.volumeLevel !== undefined) {
    slide.value = data.volumeLevel;
    // Update the volume in the background
    chrome.runtime.sendMessage({
      type: 'change-vol',
      target: 'offscreen',
      data: data.volumeLevel
    });
  } else {
    // Use default value if nothing is saved
    slide.value = DEFAULT_VOLUME;
  }
});

slide.onchange = function() {
  // Save the volume value when changed
  chrome.storage.local.set({'volumeLevel': this.value});
  
  chrome.runtime.sendMessage({
    type: 'change-vol',
    target: 'offscreen',
    data: this.value
  });
}

// --- Review prompt logic ---
const REVIEW_URL = 'https://chromewebstore.google.com/detail/volume-booster/ejkiikneibegknkgimmihdpcbcedgmpo/reviews';

chrome.storage.local.get(['useCount', 'reviewDismissed'], function(data) {
  var count = (data.useCount || 0) + 1;
  chrome.storage.local.set({ useCount: count });

  if (count >= 3 && !data.reviewDismissed) {
    document.getElementById('review-banner').style.display = 'block';
  }
});

document.getElementById('review-link').onclick = function(e) {
  e.preventDefault();
  chrome.tabs.create({ url: REVIEW_URL });
  chrome.storage.local.set({ reviewDismissed: true });
  document.getElementById('review-banner').style.display = 'none';
};

document.getElementById('review-dismiss').onclick = function(e) {
  e.preventDefault();
  chrome.storage.local.set({ reviewDismissed: true });
  document.getElementById('review-banner').style.display = 'none';
};

button.onclick = function() {
  // Reset volume to default when turning off
  chrome.storage.local.set({'volumeLevel': DEFAULT_VOLUME});
  
  chrome.runtime.sendMessage({
    type: 'stop-streaming',
    target: 'offscreen'
  });
  
  window.close();
}