/**
 * D365 Entity Compare - Content Script
 *
 * Injected into every D365 F&O tab by the manifest content_scripts declaration.
 * The popup sends a chrome.tabs.sendMessage({ type: 'FETCH_ENTITIES', endpoint })
 * and this script performs the fetch INSIDE the D365 tab using the browser's
 * existing session cookies — no bearer token required.
 */

chrome.runtime.onMessage.addListener(function(msg, _sender, sendResponse) {

  // ── DIAGNOSE: return tab URL + raw response preview for a given endpoint ──
  if (msg.type === 'DIAGNOSE') {
    var results = { tabUrl: window.location.href, endpoint: msg.endpoint, status: null, body: '', error: null };
    fetch(msg.endpoint, {
      method: 'GET',
      credentials: 'include',
      headers: { 'Accept': 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' }
    })
    .then(function(r) {
      results.status = r.status;
      return r.text();
    })
    .then(function(body) {
      results.body = body.slice(0, 800);
      sendResponse(results);
    })
    .catch(function(e) {
      results.error = e.message;
      sendResponse(results);
    });
    return true;
  }

  // ── FETCH_ENTITIES ──
  if (msg.type !== 'FETCH_ENTITIES') return false;

  var endpoint = msg.endpoint;

  fetch(endpoint, {
    method: 'GET',
    credentials: 'include',
    headers: {
      'Accept': 'application/json',
      'OData-MaxVersion': '4.0',
      'OData-Version': '4.0'
    }
  })
  .then(function(r) {
    if (!r.ok) {
      return r.text().then(function(body) {
        var detail = '';
        try {
          var j = JSON.parse(body);
          detail = (j.error && (j.error.message || j.error.code)) || '';
        } catch(e) {
          detail = body.replace(/<[^>]+>/g, ' ').replace(/\s+/g,' ').trim().slice(0, 500);
        }
        sendResponse({ ok: false, status: r.status, err: 'HTTP ' + r.status, detail: detail });
      });
    }
    return r.json().then(function(j) {
      // /Metadata/DataEntities returns a plain array []
      // /data/DataEntities      returns { value: [] }
      // Pass the raw parsed JSON — popup.js normaliseEntities() handles both shapes
      sendResponse({ ok: true, data: j });
    });
  })
  .catch(function(e) {
    sendResponse({ ok: false, err: String(e.message) });
  });

  return true;
});
