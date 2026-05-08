// Open the popup as a full-tab page when the extension icon is clicked
// while the user holds Shift, OR when a message 'OPEN_FULL_TAB' is received.
chrome.runtime.onMessage.addListener(function(msg) {
  if (msg && msg.type === 'OPEN_FULL_TAB') {
    var url = chrome.runtime.getURL('popup.html') + '?fullpage=1';
    chrome.tabs.create({ url: url });
  }
});
