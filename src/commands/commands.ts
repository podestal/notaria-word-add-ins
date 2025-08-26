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
  // No need for Word.run here since you don't use 'context'
  Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 }, function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const file = result.value;
      const sliceCount = file.sliceCount;
      let slicesReceived = 0, docdata = [];

      // Use a function expression instead of a function declaration
      const getSliceAsync = (sliceIndex: number) => {
        file.getSliceAsync(sliceIndex, function (sliceResult) {
          if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
            docdata[sliceIndex] = sliceResult.value.data;
            slicesReceived++;
            if (slicesReceived === sliceCount) {
              // All slices received, combine into a Blob
              const byteArray = docdata.reduce((acc, curr) => acc.concat(Array.from(curr)), []);
              const blob = new Blob([new Uint8Array(byteArray)], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });

              Word.run(async context2 => {
                // Get the actual filename from the document URL
                let fileName = "document.docx"; // fallback
                try {
                  const documentUrl = Office.context.document.url;
                  if (documentUrl) {
                    const urlParts = documentUrl.split('/');
                    const lastPart = urlParts[urlParts.length - 1];
                    if (lastPart && lastPart.includes('.')) {
                      fileName = decodeURIComponent(lastPart);
                    }
                  }
                  if (fileName === "document.docx") {
                    const props = context2.document.properties;
                    props.load("title");
                    await context2.sync();
                    if (props.title) {
                      fileName = props.title;
                    }
                  }
                  console.log("Extracted filename:", fileName);
                } catch (error) {
                  console.log("Error getting filename:", error);
                }
                if (!fileName.endsWith(".docx")) fileName += ".docx";

                // 2. Send to backend as FormData
                const formData = new FormData();
                formData.append("file", blob, fileName);

                // --- Dialog API integration ---
                let waitingDialog = null;

                // Open waiting dialog
                Office.context.ui.displayDialogAsync(
                  window.location.origin + "/assets/feedback-waiting.html",
                  { height: 30, width: 20, displayInIframe: true },
                  function (asyncResult) {
                    waitingDialog = asyncResult.value;

                    axios.post(`${process.env.API_BASE_URL}${process.env.API_ENDPOINT}`, formData, {
                      headers: { 
                        'Content-Type': 'multipart/formats-data',
                        'Authorization': `Token ${process.env.API_TOKEN}`
                      }
                    })
                      .then(response => {
                        console.log(response.data);
                        if (waitingDialog) waitingDialog.close();
                        Office.context.ui.displayDialogAsync(
                          window.location.origin + "/assets/feedback-success.html",
                          { height: 30, width: 20, displayInIframe: true }
                        );
                      })
                      .catch(error => {
                        console.log(error);
                        if (waitingDialog) waitingDialog.close();
                        Office.context.ui.displayDialogAsync(
                          window.location.origin + "/assets/feedback-error.html",
                          { height: 30, width: 20, displayInIframe: true }
                        );
                      })
                      .finally(() => {
                        file.closeAsync();
                        event.completed();
                      });
                  }
                );
                // --- End Dialog API integration ---
              });
            } else {
              getSliceAsync(sliceIndex + 1);
            }
          } else {
            file.closeAsync();
            event.completed();
          }
        });
      };
      getSliceAsync(0);
    } else {
      event.completed();
    }
  });
}

// Associate command functions
Office.actions.associate("saveToServer", saveToServer);
Office.actions.associate("action", action);