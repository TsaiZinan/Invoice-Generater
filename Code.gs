// Generate single invoice by id
function byNumberSingle() {
  // get Generate data from sheet
  const mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Generate");
  const invoiceNumber = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Generate").getRange(4,2,1,1).getDisplayValues()[0];
  console.log(invoiceNumber);

  errors = genBy(2, null, invoiceNumber, invoiceNumber);

  mainSheet.getRange('H3').setValue(errors);
}

// Generate multiple invoices by id
function byNumberMultiple() {
  // get Generate data from sheet
  const mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Generate");
  const invoiceNumberOne = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Generate").getRange(4,5,1,1).getDisplayValues()[0];
  const invoiceNumberTwo = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Generate").getRange(4,6,1,1).getDisplayValues()[0];
  console.log(Math.max(invoiceNumberOne, invoiceNumberTwo));
  console.log(Math.min(invoiceNumberOne, invoiceNumberTwo));

  errors = genBy(2, null, Math.min(invoiceNumberOne, invoiceNumberTwo), Math.max(invoiceNumberOne, invoiceNumberTwo)).join();
  console.log(errors);

  mainSheet.getRange('H7').setValue(errors);
}

// Generate all the invoices
function byAll() {
  // get Generate data from sheet
  const mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Generate");
  const invoiceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Test");
  const invoiceArray = invoiceSheet.getRange(2,1,invoiceSheet.getLastRow()-1,1).getDisplayValues();
  var invoiceNumber = invoiceArray[invoiceArray.length - 1][0]
  console.log(invoiceNumber);

  errors = genBy(2, null, 1, invoiceNumber).join();

  mainSheet.getRange('H9').setValue(errors);
}

// Generate all the invoices with specific name
function byName() {
  // get Generate data from sheet
  const mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Generate");
  const invoiceName = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Generate").getRange(4,4,1,1).getDisplayValues()[0].toString();
  console.log(invoiceName);

  errors = genBy(1, invoiceName).join();

  mainSheet.getRange('H5').setValue(errors);
}

//======================================================================================================================================================

//mode 1: generate by name
//mode 2: generate by number
function genBy(mode, name, min, max) {
  const docFile = DriveApp.getFileById("==============your temple doc file id==============");
  const tempFolder = DriveApp.getFolderById("==============your temporary folder id==============");
  const pdfFolder = DriveApp.getFolderById("==============your output PDF folder id==============");

  const currentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Test");
  var data = currentSheet.getRange(2,1,currentSheet.getLastRow()-1,currentSheet.getLastColumn()).getDisplayValues();

  // console.log(data);

  // get FROM data from sheet
  const mySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("From");
  const myData = mySheet.getRange(2,1,1,5).getDisplayValues();

  // get TO data from sheet
  const customerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("To");
  const customerData = customerSheet.getRange(2,1,currentSheet.getLastRow(),5).getDisplayValues();

  // console.log(data);

  //mode switch
  if (mode == 1) {
    // console.log(name);
    for( var i = 0; i < data.length; i++){ 
      if ( data[i][1] !== name) { 
        data.splice(i, 1); 
        i = i - 1
      }
    }
  } 
  else if (mode == 2) {
    // console.log(min + max)
    data = data.slice(min-1, max);
  }

  console.log(data);

  let errors = [];
  data.forEach(row => {

      try{
        // get the result of matching customer data
        let customerName = row[1];
        let matchCustomerData = customerMatch(customerData, customerName);

        createInvoicePDF(row, myData[0], matchCustomerData, docFile, tempFolder, pdfFolder);
        errors.push([row[0] + " Done!"])
      } catch(err){
        errors.push([row[0] + " Failed!"])
      }
    
  });

  return errors;
}

//match the name of customer in To sheet and return that line of result
function customerMatch(customerData, matchName) {
  let matchResult = [];
  for (index = 0; index < customerData.length; index++) {
    if (customerData[index][0] === matchName) {
      matchResult = customerData[index];
      break;
    }
  }
  return matchResult;
}

// ================= Data structure =================
// data:         [number, name, date, item, qty, total]
// myData:       [myname, myabn, myphone, myaddress_1, myaddress_2]
// customerData: [customername, customerabn, customerphone, customeraddress_1, customeraddress_2]

function createInvoicePDF(data, myData, customerData, docFile,tempFolder,pdfFolder) {

  // create tempFile doc copy from Template doc into TEMP folder
  const tempFile = docFile.makeCopy(tempFolder);

  // open the cope doc above
  const tempDocFile = DocumentApp.openById(tempFile.getId());

  // open the content
  const tempDocFileContent = tempDocFile.getBody();

  // replace the text with variable
  tempDocFileContent.replaceText("{name}", data[1]);
  tempDocFileContent.replaceText("{date}", data[2]);
  tempDocFileContent.replaceText("{item}", data[3]);
  tempDocFileContent.replaceText("{qty}", data[4]);
  tempDocFileContent.replaceText("{total}", data[5]);

  tempDocFileContent.replaceText("{myname}", myData[0]);
  tempDocFileContent.replaceText("{myabn}", myData[1]);
  tempDocFileContent.replaceText("{myphone}", myData[2]);
  tempDocFileContent.replaceText("{myaddress_1}", myData[3]);
  tempDocFileContent.replaceText("{myaddress_2}", myData[4]);

  tempDocFileContent.replaceText("{customername}", customerData[0]);
  tempDocFileContent.replaceText("{customerabn}", customerData[1]);
  tempDocFileContent.replaceText("{customerphone}", customerData[2]);
  tempDocFileContent.replaceText("{customeraddress_1}", customerData[3]);
  tempDocFileContent.replaceText("{customeraddress_2}", customerData[4]);

  // save
  tempDocFile.saveAndClose();

  // convert to PDF
  const tempPDF = tempFile.getAs(MimeType.PDF);

  // save the PDF into PDF folder and change the name of PDF with name
  pdfFolder.createFile(tempPDF).setName(data[1]+" "+data[2])

  // remove the file in TEMP folder
  // tempFolder.removeFile(tempFile);
  tempFile.setTrashed(true);

  console.log(data[1] + " " + data[2] + " done!")
}

