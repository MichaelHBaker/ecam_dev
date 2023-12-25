/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */

Office.onReady(info => {
  // Office is now ready
  if (info.host === Office.HostType.Excel) {
      // Assign event handlers and other initialization logic
  }
});

function OnAction_ECAM(event) {
const elementId = event.source['id']; // Get the button ID from the event object

Office.context.ui.displayDialogAsync('https://localhost:3000/popup.html', {height: 30, width: 20}, 
    function (asyncResult) {
        if (asyncResult.status === "failed") {
            console.error("Error displaying dialog: " + asyncResult.error.message);
            event.completed();
            return;
        }
        let dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function(arg) {
            const messageFromDialog = JSON.parse(arg.message);
            if (messageFromDialog.status === 'ok') {
                dialog.close();
                event.completed(); // Call event.completed() after the dialog is closed
            }
        });

        // Send the element ID to the dialog
        const messageToSend = JSON.stringify({ elementId: elementId });
        console.log("Sending message to dialog:", messageToSend);
        dialog.messageChild(messageToSend);
    }
);
}