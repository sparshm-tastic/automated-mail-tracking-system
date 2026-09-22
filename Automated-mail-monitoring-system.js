function onFormSubmit(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = e.range.getRow();
  
  const fileUrlColumn = 2; 
  const fileUrls = sheet.getRange(row, fileUrlColumn).getValue();
  
  if (!fileUrls) return; 

  const fileId = extractFileId(fileUrls);
  if (!fileId) return;

  try {

    const imageFile = DriveApp.getFileById(fileId);
    const resource = {
      title: imageFile.getName() + ' (OCR Temp)',
      mimeType: MimeType.GOOGLE_DOCS 
    };
    
    const tempDoc = Drive.Files.copy(resource, fileId);
    const doc = DocumentApp.openById(tempDoc.id);
    const textContent = doc.getBody().getText();
    
    DriveApp.getFileById(tempDoc.id).setTrashed(true);

    const textOutputColumn = 18; 
    sheet.getRange(row, textOutputColumn).setValue(textContent);

    
    const senderAddressMatch = textContent.match(/(?:From|Sender(?:'s Address)?)\s*[:\-]?\s*([\s\S]*?)(?:To|Receiver|Date|Article|$)/i);
    if (senderAddressMatch) {
      sheet.getRange(row, 7).setValue(senderAddressMatch[1].trim());
    }

    
    const srNoMatch = textContent.match(/(?:SR|S\.?No\.?)\s*[:\-]?\s*(\d+)/i);
    if (srNoMatch) {
      sheet.getRange(row, 8).setValue(srNoMatch[1]);
    } else {
      const fallbackSrMatch = textContent.match(/^\s*(\d+)/);
      if (fallbackSrMatch) sheet.getRange(row, 8).setValue(fallbackSrMatch[1]);
    }

   
    const boNameMatch = textContent.match(/(?:BO Name|Branch(?: Office)?)\s*[:\-]?\s*([A-Za-z\s]+)/i);
    if (boNameMatch) {
      sheet.getRange(row, 9).setValue(boNameMatch[1].trim()); 
    }

   
    const articleMatch = textContent.match(/\b[A-Z0-9]{10,15}\b/);
    if (articleMatch) {
      sheet.getRange(row, 10).setValue(articleMatch[0]); 
    }

   
    const dateMatch = textContent.match(/\b\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}\b/);
    if (dateMatch) {
      sheet.getRange(row, 11).setValue(dateMatch[0]); 
    }

   
    const otherRemarksMatch = textContent.match(/(?:Other Remarks?|Additional Note)\s*[:\-]?\s*([\s\S]*?)(?:$)/i);
    if (otherRemarksMatch) {
      sheet.getRange(row, 12).setValue(otherRemarksMatch[1].trim());
    }

   
    const hosoMatch = textContent.match(/(?:Booking\s+)?(?:HO|SO)\s*[:\-]?\s*([A-Za-z0-9\s]+)/i);
    if (hosoMatch) {
      sheet.getRange(row, 13).setValue(hosoMatch[1].trim());
    }

    
    const recipientAddressMatch = textContent.match(/(?:To|Receiver(?:'s Address)?)\s*[:\-]?\s*([\s\S]*?)(?:From|Sender|Date|Mobile|Phone|$)/i);
    if (recipientAddressMatch) {
      sheet.getRange(row, 14).setValue(recipientAddressMatch[1].trim());
    }

   
    const mobileMatch = textContent.match(/\b(?:\+?91[\-\s]?)?[6-9]\d{9}\b/);
    if (mobileMatch) {
      sheet.getRange(row, 15).setValue(mobileMatch[0]); 
    }

    
    const postmasterMatch = textContent.match(/(?:Postmaster|BPM|SPM|Branch Postmaster)\s*[:\-]?\s*([A-Za-z\s]+)/i);
    if (postmasterMatch) {
      sheet.getRange(row, 16).setValue(postmasterMatch[1].trim());
    }

  
    const remarksMatch = textContent.match(/(?:Remarks?|Note|Comments?)\s*[:\-]?\s*([\s\S]*?)(?:$)/i);
    if (remarksMatch) {
      sheet.getRange(row, 17).setValue(remarksMatch[1].trim());
    }

  } catch (error) {
    console.error("Error processing file: " + error.toString());
    sheet.getRange(row, 18).setValue("Error extracting text."); 
  }
}

function extractFileId(urlString) {
  const firstUrl = urlString.split(',')[0].trim(); 
  const match = firstUrl.match(/id=([^&]+)/);
  return match ? match[1] : null;
}