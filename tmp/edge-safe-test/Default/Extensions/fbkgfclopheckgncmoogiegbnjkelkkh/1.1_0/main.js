
function validURL(str) {
  var pattern = new RegExp('^(https?:\\/\\/)?' + // protocol
    '((([a-z\\d]([a-z\\d-]*[a-z\\d])*)\\.)+[a-z]{2,}|' + // domain name
    '((\\d{1,3}\\.){3}\\d{1,3}))' + // OR ip (v4) address
    '(\\:\\d+)?(\\/[-a-z\\d%_.~+]*)*' + // port and path
    '(\\?[;&a-z\\d%_.~+=-]*)?' + // query string
    '(\\#[-a-z\\d_]*)?$', 'i'); // fragment locator
  return !!pattern.test(str);
}


function OnClick(info, tab) {
  try {
    link = atob(String(info.selectionText).trim())
    console.log(link)
    if (validURL(link)) {
      chrome.tabs.create({ url: link });
    }
  }
  catch (e) {

  }
}

function copy2clipboard (str) {
  document.oncopy = function(event) {
    event.clipboardData.setData('text/plain', String(str));
    event.preventDefault();
  };
  document.execCommand("copy", false, null);
}

function OnClick2(info, tab) {
  try {
    ss = atob(String(info.selectionText).trim())
    copy2clipboard(ss)
    
  }
  catch (e) {

  }
}


chrome.contextMenus.create({
  "title": 'Go to base64 decoded link...',
  "contexts": ["selection"],
  "onclick": OnClick
})

chrome.contextMenus.create({
  "title": 'Copy base64 decoded string...',
  "contexts": ["selection"],
  "onclick": OnClick2
})
