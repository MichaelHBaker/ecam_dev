/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */

Office.onReady(() => {
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, function (arg) {
        const messageFromParent = JSON.parse(arg.message);
        console.log("Receiving message in dialog:", messageFromParent);


        document.getElementById("elementIdDisplay").innerText = "Button clicked: " + messageFromParent.elementId;
    });

    document.getElementById("okButton").addEventListener("click", function() {
        Office.context.ui.messageParent(JSON.stringify({status: 'ok'}));
    });
});