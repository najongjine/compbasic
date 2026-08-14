
var gainNode;
var os;
var audioCtx;
var streamer;
var source;
var prevFullScreen = false;
var val = 2;
chrome.action.setPopup({ popup: "popup.html"});
//EACH TAB CAN HAVE A CONTEXT

chrome.runtime.getPlatformInfo(function(info) {
    // Display host OS in the console
    console.error(info.os);
    os = info.os;
});


//chrome.runtime.setUninstallURL("https://jointoucan.com/partners/volumebooster", function(){
//    console.log("uninstalled");
//})


// NEW SHIT HERE

chrome.runtime.onConnect.addListener(async (tab) => {
    // Track fullscreen state
    let prevFullScreen = false;
    let os = navigator.platform.toLowerCase().includes('mac') ? 'mac' : 'win';

    // Helper function to request fullscreen
    const requestFullscreen = async (tabId) => {
        try {
            await chrome.scripting.executeScript({
                target: { tabId: tabId },
                function: () => {
                    if (!document.fullscreenElement) {
                        if (document.documentElement.requestFullscreen) {
                            document.documentElement.requestFullscreen();
                        } else if (document.documentElement.webkitRequestFullscreen) {
                            document.documentElement.webkitRequestFullscreen();
                        }
                    }
                }
            });
        } catch (err) {
            console.error('Fullscreen request failed:', err);
        }
    };

    // Message listener
    chrome.runtime.onMessage.addListener(function(msg) {
        if (msg.action === 'give value'){
            port.postMessage(val);
            console.error('sent val'+val);
        }
        if (msg.action === 'start') {
            console.error("prev sound level = "+val);
        }
        if(isNumeric(msg.action)){
            console.error("adjust volume");
            val = msg.action;
            console.error("new sound level = "+val)
            gainNode.gain.value = 2**(val);
        }
        if(off(msg.faction)){
            console.error("OFF");
        }
        // Add fullscreen request handling
        if(msg.action === 'request-fullscreen') {
            requestFullscreen(tab.id);
        }
    });

    // Status change listener with enhanced fullscreen handling
    chrome.tabCapture.onStatusChanged.addListener(function (info) {
        if(info.fullscreen) {
            if(!prevFullScreen) {
                if(os === 'mac') {
                    chrome.notifications.create({
                        type: 'basic',
                        iconUrl: 'icon48.png',
                        title: 'Fullscreen Mode',
                        message: 'Press Command + Shift + F to enter fullscreen mode',
                        priority: 2
                    });
                } else {
                    chrome.notifications.create({
                        type: 'basic',
                        iconUrl: 'icon48.png',
                        title: 'Fullscreen Mode',
                        message: 'Press F11 to enter fullscreen mode',
                        priority: 2
                    });
                }
                
                // Inject content script to show in-page message
                chrome.scripting.executeScript({
                    target: { tabId: tab.id },
                    function: () => {
                        const div = document.createElement('div');
                        div.style.cssText = `
                            position: fixed;
                            top: 20px;
                            left: 50%;
                            transform: translateX(-50%);
                            background: rgba(0, 0, 0, 0.8);
                            color: white;
                            padding: 15px 25px;
                            border-radius: 5px;
                            z-index: 999999;
                            font-family: Arial, sans-serif;
                            animation: fadeOut 5s forwards;
                        `;
                        div.textContent = navigator.platform.toLowerCase().includes('mac') ? 
                            'Press Command + Shift + F for fullscreen' : 
                            'Press F11 for fullscreen';
                        document.body.appendChild(div);
                        
                        // Remove the message after 5 seconds
                        setTimeout(() => div.remove(), 5000);
                    }
                });
            }
        }
        prevFullScreen = info.fullscreen;
    });


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

/*
  if (streamer) {
    chrome.runtime.sendMessage({
      type: 'stop-streaming',
      target: 'offscreen'
    });
    return;
  }
*/


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

    //full screen
    chrome.tabCapture.onStatusChanged.addListener(function (info){
                //console.error("status: "+info.status);
                //console.error("tabId: "+info.tabId);
                //console.error("fullscreen: "+info.fullscreen);

                if(info.fullscreen){
                    if(!prevFullScreen){
                        if(os === 'mac'){
                            console.error("fullmac");
                            chrome.action.openPopup();
                            alert('maximize the chrome window. then press cmd+shift+f');
                        }
                        if(os === 'win'){
                            console.error("fullwin");
                            alert('press F11 for fullscreen');
                        }
                    }
                }
                prevFullScreen = info.fullscreen;
            });

});



//END NEW SHIT
/*
chrome.runtime.onConnect.addListener(function(port) {
    port.onMessage.addListener(function(msg) {
            if (msg.action === 'give value'){
                port.postMessage(val);
                console.log('sent val'+val);
            }
            if (msg.action === 'start') {
                console.log("prev sound level = "+val);
                if(!streamer){
//                    audioCtx = new AudioContext();
                
                    chrome.tabCapture.capture({
                            audio : true,
                            video : false
                        }, function(stream) {
                                streamer = stream;
                                source = audioCtx.createMediaStreamSource(streamer);
                                // Create a gain node.
                                gainNode = audioCtx.createGain();
                                // Connect the source to the gain node.
                                source.connect(gainNode);
                                // Connect the gain node to the destination.
                                gainNode.connect(audioCtx.destination);
                                gainNode.gain.value = 4;

                            });
                    }
              }



        if(isNumeric(msg.action)){
                console.log("adjust volume");
                val = msg.action;
                console.log("new sound level = "+val)
                gainNode.gain.value = 2**(val);
            }

        if(off(msg.faction)){
                streamer.getAudioTracks()[0].stop();
                streamer = null;
                audioCtx.close();
                console.log("close");
                val = 2;
            }


            chrome.tabCapture.onStatusChanged.addListener(function (info){
                //console.log("status: "+info.status);
                //console.log("tabId: "+info.tabId);
                //console.log("fullscreen: "+info.fullscreen);

                if(info.fullscreen){
                    if(!prevFullScreen){
                        if(os === 'mac'){
                            console.log("fullmac");
                            alert('maximize the chrome window. then press cmd+shift+f');
                        }
                        if(os === 'win'){
                            console.log("fullwin");
                            alert('press F11 for fullscreen');
                        }
                    }
                }
                prevFullScreen = info.fullscreen;
            });

    });
});

function isNumeric(n) {
  return !isNaN(parseFloat(n)) && isFinite(n);
}

function off(m) {
    return m;
}
*/

/*
// Background script
chrome.runtime.onConnect.addListener(function(port) {
    port.onMessage.addListener(function(msg) {
        if (msg.action === 'startCapture') {
            chrome.tabCapture.capture({audio: true, video: false}, function(stream) {
                if (stream) {
                    var tabId = msg.tabId; // Make sure you pass the tab ID in the message
                    chrome.tabs.sendMessage(tabId, {action: 'start', streamId: stream.id});
                } else {
                    console.error('Error capturing tab audio stream');
                }
            });
        }
    });
});
*/