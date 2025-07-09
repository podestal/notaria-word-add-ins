/* global Office, Word */

Office.onReady(() => {
  // Office.js is ready
});

/**
 * Example of a command function triggered from the ribbon.
 */
function action(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  Office.context.mailbox?.item?.notificationMessages.replaceAsync(
    "ActionPerformanceNotification",
    message
  );

  event.completed();
}

/**
 * Save document content to a remote server.
 */
function saveToServer(event: Office.AddinCommands.Event) {
  console.log("[WORD-ADDIN] Save to Server button clicked at:", new Date().toISOString());
  Office.context.document.setSelectedDataAsync(`Button clicked!`, { coercionType: Office.CoercionType.Text });
  // Show a visible notification to confirm button click
  Word.run(async context => {
    const props = context.document.properties;
    props.load("title");
    await context.sync();

    const fileName = props.title;
    console.log("Document title:", fileName);
    alert("Document title: " + fileName);
    Office.context.document.setSelectedDataAsync(`Button clicked! ${fileName}`, { coercionType: Office.CoercionType.Text });

    Office.context.document.setSelectedDataAsync(`${fileName}`, { coercionType: Office.CoercionType.Text });
  }).catch(error => {
    console.error("[WORD-ADDIN] Error saving document:", error);
    alert("Error: " + error.message);
  }).finally(() => {
    console.log("[WORD-ADDIN] Function completed");
    event.completed();
  });
}

// Associate command functions
Office.actions.associate("saveToServer", saveToServer);
Office.actions.associate("action", action);