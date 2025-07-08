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
  Word.run(async context => {
    const body = context.document.body;
    body.load("text");
    await context.sync();

    const documentText = body.text;

    // 🔁 Replace this with your actual API URL
    await fetch("https://yourserver.com/api/save", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ content: documentText })
    });

    console.log("Document content sent to server");
  }).catch(error => {
    console.error("Error saving document:", error);
  }).finally(() => {
    event.completed();
  });
}

// Associate command functions
Office.actions.associate("saveToServer", saveToServer);
Office.actions.associate("action", action);
