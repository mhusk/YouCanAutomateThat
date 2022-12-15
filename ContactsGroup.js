function onOpen(){
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
    .addItem('Create a Group', 'createGroup')
    .addItem('Add Contacts to Group', 'addContactsToGroup')
    .addToUi()
}

function createGroup(){
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Enter the name of the group: ');
  if(response.getSelectedButton() == ui.Button.OK){
    const groupName = response.getResponseText();
    ContactsApp.createContactGroup(groupName);
  }
}

function addContactsToGroup(){
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Add Contacts to Group','Enter the name of the group you want to add these contacts to:',ui.ButtonSet.OK_CANCEL)
  if(response.getSelectedButton() == ui.Button.OK){
    const contactGroup = doesGroupExist(response.getResponseText())
    if(contactGroup !== false){
      selectContactData().forEach(contact => {
        let newContact = ContactsApp.createContact(contact.firstName, contact.lastName,contact.lastName);
        newContact.addCompany(contact.company, contact.jobTitle)
        newContact.addToGroup(contactGroup)
        SpreadsheetApp.getActiveSheet().getRange(contact.sheetRow,6).setValue(newContact.getId());
      })
    } else{
      ui.alert('Group Does Not Exist',`List of current groups: ${ContactsApp.getContactGroups().map(group =>` ${group.getName()}` )}`,ui.ButtonSet.OK)
    }
  }
}

function doesGroupExist(input){
  let currentGroups = ContactsApp.getContactGroups().map(group => group.getName().toUpperCase());
  let indexOf = currentGroups.indexOf(input.toUpperCase());
  if(indexOf !== -1){
      let contactGroup = ContactsApp.getContactGroups()[indexOf];
      return contactGroup
  } else{
    return false
  }
}

function selectContactData(){
  const sheetData = SpreadsheetApp.getActiveSheet().getDataRange().getValues();
  const headerRow = sheetData[0];
  let output = [];
  sheetData.forEach((row, i) => {
    if(i > 0){
      output.push({
        'firstName': row[0],
        'lastName': row[1],
        'email': row[2],
        'company': row[3],
        'jobTitle': row[4],
        'sheetRow': i+1
      })
    }
  })
  return output
}
