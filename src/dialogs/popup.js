/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */
// Office.onReady((info) => {
//     // document.getElementById("ok-button").onclick = () => tryCatch(sendStringToParentPage);
// });
Office.onReady(() => {
    // If needed, Office.js is ready to be called
    console.log("popup js office onready");
  });
  
// function sendStringToParentPage() {
//     const userName = document.getElementById("name-box").value;
//     Office.context.ui.messageParent(userName);
// }

// /** Default helper for invoking an action and handling errors. */
// async function tryCatch(callback) {
//     try {
//         await callback();
//     } catch (error) {
//         // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
//         console.error(error);
//     }
// }