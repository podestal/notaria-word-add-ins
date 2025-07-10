import axios from "axios";

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
    const customProps = context.document.properties.customProperties;
    customProps.load("items");
    await context.sync();
  
    let docId = null;
    for (const prop of customProps.items) {
      if (prop.key === "documentoGeneradoId") {
        docId = prop.value;
        break;
      }
    }
    console.log("Document ID:", docId);
    axios.post("127.0.0.1:8001/docs/documentos/", {
      Kardex: docId,
    }).then(response => {
      console.log("Response:", response.data);
    }).catch(error => {
      console.error("Error:", error);
    });
  });
}

// Associate command functions
Office.actions.associate("saveToServer", saveToServer);
Office.actions.associate("action", action);