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
  
  // Show a visible notification to confirm button click
  Office.context.document.setSelectedDataAsync("Button clicked! Processing...", { coercionType: Office.CoercionType.Text });
  
  Word.run(async context => {
    console.log("[WORD-ADDIN] Starting to read document content...");
    
    const body = context.document.body;
    body.load("text");
    await context.sync();
    
    const documentText = body.text;
    console.log("[WORD-ADDIN] Document text length:", documentText.length);
    console.log("[WORD-ADDIN] Document preview:", documentText.substring(0, 100) + "...");

    // 🔁 Replace this with your actual API URL
    console.log("[WORD-ADDIN] Attempting to send to server...");
    const response = await fetch("https://yourserver.com/api/save", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ content: documentText })
    });

    console.log("[WORD-ADDIN] Server response status:", response.status);
    console.log("[WORD-ADDIN] Document content sent to server successfully");
  }).catch(error => {
    console.error("[WORD-ADDIN] Error saving document:", error);
  }).finally(() => {
    console.log("[WORD-ADDIN] Function completed");
    event.completed();
  });
}

// Associate command functions
Office.actions.associate("saveToServer", saveToServer);
Office.actions.associate("action", action);
