/**
 * Test function to fetch and log historical volume data from Yahoo Finance.
 * Useful when GOOGLEFINANCE volume feeds experience backend delays.
 */
function testGetYahooVolume() {
  Trns.trs()
  var ticker = "AVGO";
  
  // Set date range (YYYY-MM-DD)
  var startDate = "2026-09-01";
  var endDate   = "2026-09-10";
  
  Logger.log("Fetching Yahoo Finance volume data for: " + ticker);
  Logger.log("Period: " + startDate + " to " + endDate);
  
  var volumeData = getYahooVolume(ticker, startDate, endDate);
  
  Logger.log("--- RESULTS ---");
  if (volumeData && volumeData.length > 0) {
    volumeData.forEach(function(row) {
      Logger.log("Date: " + row.date + " | Volume: " + row.volume.toLocaleString());
    });
  } else {
    Logger.log("⚠️ No volume data returned or request failed.");
  }
}

/**
 * Core engine to download and parse historical volume from Yahoo Finance CSV stream.
 * 
 * @param {string} ticker Stock symbol (e.g., "AVGO")
 * @param {string|Date} startDate Start date
 * @param {string|Date} endDate End date
 * @return {Array<Object>} Array of objects containing { date, volume }
 */
function getYahooVolume(ticker, startDate, endDate) {
  try {
    // Convert dates to Unix Timestamps (in seconds)
    var p1 = Math.floor(new Date(startDate).getTime() / 1000);
    var p2 = Math.floor(new Date(endDate).getTime() / 1000);
    
    // Yahoo Finance CSV download URL format
    var url = "https://query1.finance.yahoo.com/v7/finance/download/" + encodeURIComponent(ticker) +
              "?period1=" + p1 +
              "&period2=" + p2 +
              "&interval=1d&events=history&includeAdjustedClose=true";
    
    var options = {
      "muteHttpExceptions": true,
      "headers": {
        // User-Agent helps prevent request blocking by Yahoo's endpoint
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
      }
    };
    
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      Logger.log("❌ HTTP Error: " + responseCode + " - " + response.getContentText());
      return null;
    }
    
    var csvText = response.getContentText();
    var lines = csvText.split("\n");
    var results = [];
    
    // Parse CSV rows (Skip header row 0)
    for (var i = 1; i < lines.length; i++) {
      var line = lines[i].trim();
      if (!line) continue;
      
      var cols = line.split(",");
      // CSV format: Date[0], Open[1], High[2], Low[3], Close[4], Adj Close[5], Volume[6]
      var dateStr = cols[0];
      var rawVolume = cols[6];
      
      if (dateStr && rawVolume && !isNaN(Number(rawVolume))) {
        results.push({
          date: dateStr,
          volume: Number(rawVolume)
        });
      }
    }
    
    return results;
    
  } catch (error) {
    Logger.log("❌ Execution Error: " + error.toString());
    return null;
  }
}
/**
 * Retrieves a valid Cookie and Crumb pair from Yahoo Finance.
 * Returns an object containing { cookie: string, crumb: string }
 */
function getYahooSession() {
  try {
    // Step 1: Initial request to Yahoo to capture session cookies
    var initUrl = "https://fc.yahoo.com";
    var initOptions = {
      "muteHttpExceptions": true,
      "headers": {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
      }
    };

    var initResponse = UrlFetchApp.fetch(initUrl, initOptions);
    var headers = initResponse.getAllHeaders();
    
    // Extract Set-Cookie header
    var setCookieHeader = headers["Set-Cookie"] || headers["set-cookie"];
    if (!setCookieHeader) {
      Logger.log("❌ Failed to retrieve Set-Cookie header.");
      return null;
    }

    // Format cookies into a single header string
    var cookieStr = "";
    if (Array.isArray(setCookieHeader)) {
      cookieStr = setCookieHeader.map(function(c) { return c.split(';')[0]; }).join('; ');
    } else {
      cookieStr = setCookieHeader.split(';')[0];
    }

    // Step 2: Fetch the Crumb using the active cookie session
    var crumbUrl = "https://query2.finance.yahoo.com/v1/test/getcrumb";
    var crumbOptions = {
      "muteHttpExceptions": true,
      "headers": {
        "Cookie": cookieStr,
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
      }
    };

    var crumbResponse = UrlFetchApp.fetch(crumbUrl, crumbOptions);
    var crumb = crumbResponse.getContentText().trim();

    if (!crumb || crumbResponse.getResponseCode() !== 200) {
      Logger.log("❌ Failed to retrieve crumb. Response: " + crumb);
      return null;
    }

    Logger.log("✅ Session Authenticated Successfully!");
    Logger.log("Crumb: " + crumb);

    return {
      cookie: cookieStr,
      crumb: crumb
    };

  } catch (e) {
    Logger.log("❌ Session Error: " + e.toString());
    return null;
  }
}

/**
 * Example usage: Authenticated request to fetch ticker historical data
 */
function testAuthenticatedFetch() {
  var session = getYahooSession();
  if (!session) return;

  var ticker = "AVGO";
  var url = "https://query2.finance.yahoo.com/v7/finance/download/" + ticker + 
            "?period1=1725148800&period2=1725926400&interval=1d&events=history&crumb=" + session.crumb;

  var options = {
    "muteHttpExceptions": true,
    "headers": {
      "Cookie": session.cookie,
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
  };

  var response = UrlFetchApp.fetch(url, options);
  Logger.log("HTTP Code: " + response.getResponseCode());
  Logger.log("Content:\n" + response.getContentText().substring(0, 300));
}