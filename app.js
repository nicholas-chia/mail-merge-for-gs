// create custom menu to run script on spreadsheet open
const onInstall = (e) => {
  onOpen(e)
}

const onOpen = (e) => {
  SpreadsheetApp.getUi()
  .createAddonMenu()
  .addItem('Initialise Template Sheet', 'initTemplate')
  .addItem('Initialise Configurations Sheet', 'initConfig')
  .addSeparator()
  .addItem('Notify By Email', 'confirmNotifyByEmail')
  .addToUi()
}

const showMessage = () => {
  const ui = SpreadsheetApp.getUi()
  const buttonpressed = ui.alert("Are you sure?", ui.ButtonSet.OK_CANCEL)
  if (buttonpressed == ui.Button.OK) {
    return(true)
  } else {
    return (null)
  }
}

const confirmNotifyByEmail = () => {
  const go = showMessage()
  if (go) {
    notifyByEmail()
  }
}

const notifyByEmail = () => {
  sendEmails()
}