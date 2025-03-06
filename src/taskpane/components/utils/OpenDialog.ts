export const OpenDialog = () => {
  let dialog: Office.Dialog; // Declare dialog globally for further use.

  // Determine the URL based on the environment.
//   const dialogUrl =
//     process.env.NODE_ENV === "production"
//       ? "https://shahzadumar-w.github.io/Word_Addin_Ai_DEtector/ReactDialog.html"
//       : "https://localhost:3000/ReactDialog.html";
let dialogUrl='https://localhost:3000/ReactDialog.html'
  // Open the dialog using displayDialogAsync.
  Office.context.ui.displayDialogAsync(
    dialogUrl,
    { height: 60, width: 60 },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        dialog = asyncResult.value;

        // Attach an event handler for messages received from the dialog.
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
      } else {
        console.error(`Failed to open dialog: ${asyncResult.error.message}`);
      }
    }
  );
};

// Function to process messages received from the dialog.
function processMessage(arg: Office.DialogParentMessageReceivedEventArgs) {
  try {
    const messageFromDialog = JSON.parse(arg.message); // Parse the JSON message.
    console.log("Message received from dialog:", messageFromDialog);
  } catch (error) {
    console.error("Error parsing message from dialog:", error);
  }
}
