const CACHE_KEY = 'WINDOW_STATE';

// Call this when opening window
function openWindow() {
  CacheService.getScriptCache().put(CACHE_KEY, 'OPEN', 30); // 30s expiration
  return HtmlService.createHtmlOutputFromFile('Client').setTitle('My Window');
}

// Called by client when window closes
function setWindowClosed() {
  CacheService.getScriptCache().remove(CACHE_KEY);
}

// Check window status
function isWindowOpen() {
  return CacheService.getScriptCache().get(CACHE_KEY) === 'OPEN';
}
