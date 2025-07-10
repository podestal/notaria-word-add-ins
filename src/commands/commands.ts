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
    // 1. Read custom property
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

    // 2. Write debug info to document
    context.document.body.insertParagraph(
      "DEBUG: About to POST to backend: " + docId,
      Word.InsertLocation.end
    );
    await context.sync();

    // 3. Make the POST request
    try {
      const response = await axios.post("https://quenteh.podestalservers.com/docs/documentos/", { Kardex: docId });
      // 4. Write response to document
      context.document.body.insertParagraph(
        "DEBUG: Response: " + JSON.stringify(response.data),
        Word.InsertLocation.end
      );
      await context.sync();
    } catch (error: any) {
      // 5. Write detailed error info to document
      let errorDetails = [
        "DEBUG: Error: " + error,
        "AxiosError.message: " + (error.message || ""),
        "AxiosError.code: " + (error.code || ""),
        "AxiosError.config: " + JSON.stringify(error.config || {}),
        "AxiosError.response: " + (error.response ? JSON.stringify(error.response.data) : "No response"),
        "AxiosError.request: " + (error.request ? error.request.toString() : "No request"),
        "AxiosError.stack: " + (error.stack || "")
      ].join("\n");

      context.document.body.insertParagraph(
        errorDetails,
        Word.InsertLocation.end
      );
      await context.sync();
    }
  })
  .catch(error => {
    Word.run(async context2 => {
      context2.document.body.insertParagraph(
        "DEBUG: Word.run Error: " + error,
        Word.InsertLocation.end
      );
      await context2.sync();
    });
  })
  .finally(() => {
    event.completed();
  });
}

// Associate command functions
Office.actions.associate("saveToServer", saveToServer);
Office.actions.associate("action", action);