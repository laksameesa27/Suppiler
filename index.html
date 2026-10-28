function parseVendorEmailsToSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // ค้นหาอีเมลที่ยังไม่ได้อ่านและมีคำว่า "ใบเสนอราคา" หรือ "Quotation"
  var threads = GmailApp.search('is:unread "Quotation" OR "ใบเสนอราคา"');
  
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var emailBody = messages[j][PlainBody];
      var emailDate = messages[j].getDate();
      
      // ส่งเนื้อหาอีเมลไปให้ Gemini API สกัดข้อมูลออกมาเป็น JSON
      var extractedData = callGeminiAPI(emailBody);
      
      if (extractedData) {
        // นำข้อมูลที่ได้มาหยอดลงแถวใหม่ใน Google Sheets
        sheet.appendRow([
          emailDate,
          extractedData.company_name,
          extractedData.contact_person,
          extractedData.email,
          extractedData.product_category,
          "Pending"
        ]);
      }
    }
    // ทำเครื่องหมายว่าอ่านแล้ว เพื่อไม่ให้ดึงซ้ำ
    threads[i].markRead();
  }
}

function callGeminiAPI(text) {
  var apiKey = "ใส่_GEMINI_API_KEY_ของคุณที่นี่"; 
  var url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + apiKey;
  
  var prompt = "Extract company name, contact person, email, and product category from this email text as JSON format with keys: company_name, contact_person, email, product_category:\n\n" + text;
  
  var payload = {
    "contents": [{ "parts": [{"text": prompt}] }]
  };
  
  var options = {
    "method" : "post",
    "contentType" : "application/json",
    "payload" : JSON.stringify(payload),
    "muteHttpExceptions": true
  };
  
  try {
    var response = UrlFetchApp.fetch(url, options);
    var json = JSON.parse(response.getContentText());
    var rawText = json.candidates[0].content.parts[0].text;
    
    // ทำความสะอาด JSON string ที่ได้จาก AI
    var cleanJson = rawText.replace(/```json/g, "").replace(/```/g, "").trim();
    return JSON.parse(cleanJson);
  } catch (error) {
    Logger.log("Error parsing AI response: " + error);
    return null;
  }
}
