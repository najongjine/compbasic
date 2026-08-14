var gainNode;
var os;
var audioCtx;
var streamer;
var source;
var prevFullScreen = false;
var val = 2;
//EACH TAB CAN HAVE A CONTEXT


chrome.runtime.getPlatformInfo(function (info) {
  // Display host OS in the console
  console.error(info.os);
  os = info.os;
});


chrome.runtime.onConnect.addListener(async (tab) => {
  const existingContexts = await chrome.runtime.getContexts({});
  let streamer = false;

  const offscreenDocument = existingContexts.find(
    (c) => c.contextType === 'OFFSCREEN_DOCUMENT'
  );

  // If an offscreen document is not already open, create one.
  if (!offscreenDocument) {
    // Create an offscreen document.
    await chrome.offscreen.createDocument({
      url: 'offscreen.html',
      reasons: ['USER_MEDIA'],
      justification: 'Streaming from chrome.tabCapture API'
    });
  } else {
    streamer = offscreenDocument.documentUrl.endsWith('#streaming');
  }


  // Get a MediaStream for the active tab.
  const streamId = await chrome.tabCapture.getMediaStreamId({
    targetTabId: tab.id
  });

  // Send the stream ID to the offscreen document to start streaming.
  chrome.runtime.sendMessage({
    type: 'start-streaming',
    target: 'offscreen',
    data: streamId
  });

});


// When a video in the captured tab enters fullscreen, Chrome keeps the browser
// tab strip visible as a capture indicator. Putting the WINDOW into browser
// fullscreen hides it. We remember the window's prior state so we can restore it
// exactly on exit (so we never leave the user somewhere they didn't start), and
// if the window is ALREADY fullscreen we toggle through 'maximized' to force
// Chrome to re-lay-out a chrome-less fullscreen -- otherwise the capture-induced
// tab strip stays put because there's no state transition to trigger the relayout.
var windowStateBeforeFullscreen = null;

chrome.tabCapture.onStatusChanged.addListener(function (info) {
  // Only react to an actual enter/exit transition.
  if (info.fullscreen === prevFullScreen) {
    return;
  }
  prevFullScreen = info.fullscreen;

  chrome.windows.getCurrent((win) => {
    if (chrome.runtime.lastError || !win) {
      return;
    }

    if (info.fullscreen) {
      // Video entered fullscreen: remember where we were, then go chrome-less.
      windowStateBeforeFullscreen = win.state;
      if (win.state === 'fullscreen') {
        // Already fullscreen -- force a re-layout so the tab strip hides.
        chrome.windows.update(win.id, { state: 'maximized' }, () => {
          chrome.windows.update(win.id, { state: 'fullscreen' });
        });
      } else {
        chrome.windows.update(win.id, { state: 'fullscreen' });
      }
    } else {
      // Video exited fullscreen: restore exactly the state we started from.
      const restoreTo = windowStateBeforeFullscreen || 'normal';
      windowStateBeforeFullscreen = null;
      if (win.state !== restoreTo) {
        chrome.windows.update(win.id, { state: restoreTo });
      }
    }
  });
});
