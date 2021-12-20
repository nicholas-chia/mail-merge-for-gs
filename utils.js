// const log2Sheet = (log) => {
//   const ss = SpreadsheetApp.getActiveSpreadsheet()
//   const sheet = ss.getSheetByName("Logs")
//   const lastRow = sheet.getLastRow() + 1
//   sheet.getRange(lastRow, 1).setValue(log)
//   sheet.getRange(lastRow, 2).setValue(Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'"))
// }

const addConfig = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheetByName("Configurations")
  if (sheet == null) {
    ss.insertSheet("Configurations")
    sheet = ss.getSheetByName("Configurations")
    // Each configs becomes an array [startRow, 1], [lastRowOffset, 0]   
    const configs = {
      startRow: "6",
      lastRowOffset: "0",
      banner: "https://drive.google.com/file/d/1q68Mk9RuBkMVnyLvX3C1t2ehcH24uGe4/view?usp=sharing",
      template: "https://drive.google.com/file/d/18YueNqDdP1IE6m3WEDFMZG-mGme-iv8T/view?usp=sharing",
      fromName: "APAC Enablement [Mail Merge for GS]",
      fromEmail: "nicholas.cs.chia@example.com",
      ccEmails: "nicholas.cs.chia@example.com, nicholas.cs.chia@example.com",
      attachment: "https://drive.google.com/file/d/1ko3zWBy9Js15XaG0QmH86RryyfHrORRp/view?usp=sharing"
    }
    var row = 1
    sheet.getRange(row, 1).setValue("Key").setFontWeight("bold")
    sheet.getRange(row, 2).setValue("Value").setFontWeight("bold")
    Object.entries(configs).forEach(configs => {
      row++
      // configs becomes an array [startRow, 1]
      // configs[0] is startRow and configs[1] is 
      sheet.getRange(row, 1).setValue(configs[0])
      sheet.getRange(row, 2).setValue(configs[1])
    })   
    return true
  } else {
    return false
  }
}

const rmConfig = () => {
  const ss = SpreadsheetApp.getActive()
  var sheet = ss.getSheetByName("Configurations")
  if (sheet !== null) {
    ss.deleteSheet(sheet)
  } else {
    return false
  }
}

const initConfig = () => {
  rmConfig()
  // log2Sheet("Config sheet removed.")
  addConfig()
  // log2Sheet("Default config added to sheet.")
}

const readConfigs = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getSheetByName("Configurations")
  const startRow = 2
  const keyCol = 1
  const valueCol = 2
  const numOfRows = sheet.getLastRow() - 1
  var key = sheet.getRange(startRow, keyCol, numOfRows).getValues()
  var value = sheet.getRange(startRow, valueCol, numOfRows).getValues()
  var configs = {}
  for (i = 0; i < key.length; ++i) {
    configs[key[i]] = value[i]
  }
}

const rmTemplate = () => {
  const ss = SpreadsheetApp.getActive()
  var sheet = ss.getSheetByName("Template")
  if (sheet !== null) {
    ss.deleteSheet(sheet)
  } else {
    return false
  }
}

const initTemplate = () => {
  rmTemplate()
  // log2Sheet("Template sheet removed.")
  addTemplate()
  // log2Sheet("Created Template Sheet and added sample data.")
}

const addTemplate = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheetByName("Template")
  if (sheet == null) {
    ss.insertSheet("Template")
    sheet = ss.getSheetByName("Template")
    // Each entry becomes an array [startRow, 1], [lastRowOffset, 0]   
    const course = {
      name: "Advanced Deployment with Red Hat Ansible Automation",
      startDate: "January 17, 2022",
      endDate: "January 21, 2022",
      trgSchedule: "Mon-Thurs: 8am-4pm & 8am-11am"
    }
    // Print course info
    var row = 0
     Object.entries(course).forEach(course => {
      row++
      sheet.getRange(row, 1).setValue(course[0])
      sheet.getRange(row, 2).setValue(course[1])
    })
    const headers = {
      1: "First Name",
      2: "Last Name",
      3: "Email Address",
      4: "Start Enablement Email Sent by Who & When?"
    }
    //Print headers
    row = 5
    var col = 0
    Object.entries(headers).forEach(headers => {
      col++
      sheet.getRange(row, col).setValue(headers[1]).setFontWeight("bold")
    })
    const students = {
      1: ["Nicholas1", "Chia", "nicholas.chia@example.com", "Sent by Mail Merge for GS"],
      2: ["Nicholas2", "Chia", "nicholas.chia@example.com", "Sent by Mail Merge for GS"],
      3: ["Nicholas3", "Chia", "nicholas.chia@example.com", "Sent by Mail Merge for GS"],
      4: ["Nicholas4", "Chia", "nicholas.chia@example.com", "Sent by Mail Merge for GS"],
    }
    // Print sample student info
    row = 6
    Object.entries(students).forEach(students => {
      for (col = 0; col < students[1].length; col++){
      sheet.getRange(row, col+1).setValue(students[1][col])
      }
      row++
    })
    return true
  } else {
    return false
  }
}

const getConfigs = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getSheetByName("Configurations")
  const configs = sheet.getRange(2, 1, sheet.getLastRow()-1, sheet.getLastColumn()).getValues()
  return Object.fromEntries(configs)
}

const getCourseInfo = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getActiveSheet()
  // const sheet = ss.getSheetByName("Template")
  const course = sheet.getRange(1, 1, 4, 2).getValues()
  console.log(Object.fromEntries(course))
  return Object.fromEntries(course)
}

const getStudentsFromSheet = () => {
  const configs = getConfigs()
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getActiveSheet()
  // const sheet = ss.getSheetByName("Template")
  const students = sheet.getRange(configs.startRow, 1, sheet.getLastRow()-configs.startRow+1, sheet.getLastColumn()).getValues()
  return students
}

const getFileId = (url) => {
  var fileId
  if (url.split("/")[4] === "u") {
    fileId = url.split("/")[7]
  } else {
    fileId = url.split("/")[5]
  }
  return fileId
}

const sendEmails = () => {
  const configs = getConfigs()
  const course = getCourseInfo()
  const students = getStudentsFromSheet()
  const bannerBlob = DriveApp.getFileById(getFileId(configs.banner)).getBlob().setName("banner")
  const templateBlob = DriveApp.getFileById(getFileId(configs.template)).getBlob().setName("template")
  const attachementBlob = DriveApp.getFileById(getFileId(configs.attachment)).getBlob().setName("guide")

  students.forEach(student => {
    var person = {
      firstName: student[0],
      lastName:  student[1],
      email:     student[2],
      notified:  student[3],
      startDate: course.startDate,
      endDate: course.endDate,
      courseName: course.name,
      schedule: course.trgSchedule,

    }

    if (person.notified == "") {
      var subject = " [INTERNAL] Start Your Enablement for " + person.firstName
                  + " | " + Utilities.formatDate(course.startDate, "GMT", "dd MMM yyyy")
                  + " - " + Utilities.formatDate(course.endDate, "GMT", "dd MMM yyyy")
                  + " | Session " + course.name
      
      var templ = HtmlService.createTemplate(templateBlob)
      templ.person = person
      const message = templ.evaluate().getContent()
      const body = "This email renders only in HTML. Please contact us if you are reading this message."
      const fromEmail = configs.fromEmail
      const ccEmails = configs.ccEmails
      const senderName = configs.fromName

      GmailApp.sendEmail(person.email, subject, body, {
        from: fromEmail,
        cc: ccEmails,
        name: senderName,
        attachments: [attachementBlob.getAs(MimeType.PDF)],
        htmlBody: 
          message,
          inlineImages: {
            image: bannerBlob 
          }
      })
    }
  })
  setAllNotified()
}

const setAllNotified = () => {
  const configs = getConfigs()
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getActiveSheet()
  // const sheet = ss.getSheetByName("Template")
  const message = "Sent on " + new Date() + " by Mail Merge for GS."
  for (var row = configs.startRow; row <= sheet.getLastRow()-configs.lastRowOffset; row++) {
    //console.log(sheet.getRange(row, 4).getValue())
    if (sheet.getRange(row, 4).getValue() == "") {
      sheet.getRange(row, 4).setValue(message)
    }
  }
}