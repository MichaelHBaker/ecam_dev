/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */

let dialog;
let message_from_parent;

Office.onReady((info) => {
  //info can be used to customize UI
  console.log(info.host.toString());
  console.log(info.platform.toString());
});

async function OnAction_ECAM(event) {
  var function_name;

  // Call function based on the button ID
  function_name = event.source['id'].replace(/^[a-z]+|\d+$/g, ''); //removes lower case prefix and numeric suffix

  message_from_parent = "Button (" + function_name + ") not working yet!";

  //add process message from taskpane, add a listner to taskpane and then modify taskpane based on the button id
  //create this tutorial again https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/excel-quickstart-jquery?tabs=yeomangenerator
  if (typeof window[function_name] === 'function') {
    message_from_parent = "Button clicked for (" + window[function_name]() + ")";
  } 
   
  openDialog();

  event.completed();
}

function SelectIntervalData() {
  return "SelectIntervalData";  
}

function openDialog() {
  const dialogUrl = 'https://localhost:3000/popup.html';

  Office.context.ui.displayDialogAsync(dialogUrl, { height: 10, width: 20 }, function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Failed to open dialog: " + asyncResult.error.message);
          return;
      }

      dialog = asyncResult.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessageFromDialog);

      
  });
}

function processMessageFromDialog(arg) {
  if (arg.message === "dialogReady") {
    dialog.messageChild(message_from_parent);
  } else {
      console.log("arg message:" + arg.message);
  }
}

