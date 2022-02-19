function onOpen(e) {
    let ui = SpreadsheetApp.getUi();
    ui.createMenu('ตัวเลือกเพิ่มเติม')
    .addItem('make google slide', 'createCertificates')
    .addSeparator()
  .addItem('make pdf', 'setpdffromgoogleslide')
    .addToUi();
  }
  
  
  function createCertificates() {
  
  
  let slideTemplateId = "ไอดีเทมเพลต google slide แม่แบบของท่าน";
  let tempFolderId = "ไอดีโฟลเดอร์เปล่าทีท่านต้องการเก็บ ไฟล์ google slide"; 
  
  
    let template = DriveApp.getFileById(slideTemplateId);
  
  
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    let values = sheet.getDataRange().getValues();
    let headers = values[0];
    let nameemployee = headers.indexOf("header ของคอลัมพ์ท่าน");    //1
    
  
    let SlideIndex = headers.indexOf("header slideID ของคอลัมพ์ท่าน");                      
    let statusIndex = headers.indexOf("header สำหรับค่าสถานะ ของคอลัมพ์ท่าน");                   
    
  
  
    for (let i = 1; i < values.length; i++) {
      let rowData = values[i];
      let N1 = rowData[nameemployee];
      let tempFolder = DriveApp.getFolderById(tempFolderId);
      let empSlideId = template.makeCopy(tempFolder).setName(N1).getId();        
      let empSlide = SlidesApp.openById(empSlideId).getSlides()[0];
  
  
      empSlide.replaceAllText("ตัวแปรในไฟล์ google slide ที่ท่านอยากให้เติมคำ", N1);
  
  
  
      sheet.getRange(i + 1, SlideIndex + 1).setValue(empSlideId);
      sheet.getRange(i + 1, statusIndex + 1).setValue("สร้างสำเร็จ");
      SpreadsheetApp.flush();
   }
  }
  
  
  
  
  
  
  
  
  function setpdffromgoogleslide() {
  //******************************************************* */
  
    const fid = 'ไอดีโฟลเดอร์เปล่าทีท่านต้องการเก็บ ไฟล์ pdf '; //โฟลเดอร์ที่จะเก็บ pdf โดยมันจะแยกกับ โฟลเดอร์ที่เก็บ google slide
  //************************************ ********************/
  
    const folder = DriveApp.getFolderById(fid);
  
  
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); //ดึงค่าจาก google sheet ที่กำลังรัน
    let values = sheet.getDataRange().getValues();
    let headers = values[0];
  
    let nameemployee = headers.indexOf("header ของคอลัมพ์ท่าน");  
  
  //*************************************************************** */
  
    let SlideIndex = headers.indexOf("header slideID ของคอลัมพ์ท่าน");  //จำเป็น id สไลด์ โดยจะสร้าง pdf จาก slide นั้นๆ ห้ามลบ google slide ก่อนสร้าง pdf หากสร้างเสร็จลบได้
  
  //*************************************************** */
  
   for (let i = 1; i < values.length; i++) {
     
      let rowData = values[i];
      let idpdf = rowData[SlideIndex]; // 
      let Namefile = rowData[nameemployee];
  
      Logger.log(idpdf);
    const blob = DriveApp.getFileById(idpdf).getBlob();
    const file = folder.createFile(blob);
    file.setName(Namefile); // ชื่อของไฟล์ที่จะเก็บ 
  
  }
  }
  
