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

              // Try to get the file name from the document properties
              Word.run(async context2 => {
                // Get the actual filename from the document URL
                let fileName = "document.docx"; // fallback
                
                try {
                  // Try to get the filename from the document URL
                  const documentUrl = Office.context.document.url;
                  if (documentUrl) {
                    const urlParts = documentUrl.split('/');
                    const lastPart = urlParts[urlParts.length - 1];
                    if (lastPart && lastPart.includes('.')) {
                      // Decode URL-encoded characters to preserve underscores and special chars
                      fileName = decodeURIComponent(lastPart);
                    }
                  }
                  
                  // If we couldn't get it from URL, try document properties
                  if (fileName === "document.docx") {
                    const props = context2.document.properties;
                    props.load("title");
                    await context2.sync();
                    if (props.title) {
                      fileName = props.title;
                    }
                  }
                  
                  // Debug: Log the filename we're using
                  console.log("Extracted filename:", fileName);
                  // Word.run(async contextDebug => {
                  //   contextDebug.document.body.insertParagraph(
                  //     "DEBUG: Using filename: " + fileName,
                  //     Word.InsertLocation.end
                  //   );
                  //   await contextDebug.sync();
                  // });
                } catch (error) {
                  console.log("Error getting filename:", error);
                }

                // Ensure .docx extension
                if (!fileName.endsWith(".docx")) fileName += ".docx";

                // 2. Send to backend as FormData
                const formData = new FormData();
                formData.append("file", blob, fileName);

                axios.post("https://quenteh.podestalservers.com/docs/upload-docx/", formData, {
                  headers: {
                    'Content-Type': 'multipart/form-data'
                  }
                })
                  .then(response => {
                    console.log(response.data);
                    // Word.run(async context3 => {
                    //   context3.document.body.insertParagraph(
                    //     "DEBUG: Upload response: " + JSON.stringify(response.data),
                    //     Word.InsertLocation.end
                    //   );
                    //   await context3.sync();
                    // });
                  })
                  .catch(error => {
                    console.log(error);
                    // Word.run(async context3 => {
                    //   context3.document.body.insertParagraph(
                    //     "DEBUG: Upload error: " + error,
                    //     Word.InsertLocation.end
                    //   );
                    //   await context3.sync();
                    // });
                  })
                  .finally(() => {
                    file.closeAsync();
                    event.completed();
                  });
              });
            } else {
              getSliceAsync(sliceIndex + 1);
            }
          } else {
            // Word.run(async context2 => {
            //   context2.document.body.insertParagraph(
            //     "DEBUG: Error getting slice: " + sliceResult.error.message,
            //     Word.InsertLocation.end
            //   );
            //   await context2.sync();
            // });
            file.closeAsync();
            event.completed();
          }
        });
      };
      getSliceAsync(0);
    } else {
      // Word.run(async context2 => {
      //   context2.document.body.insertParagraph(
      //     "DEBUG: Error getting file: " + result.error.message,
      //     Word.InsertLocation.end
      //   );
      //   await context2.sync();
      // });
      event.completed();
    }
  });
}

// Associate command functions
Office.actions.associate("saveToServer", saveToServer);
Office.actions.associate("action", action);