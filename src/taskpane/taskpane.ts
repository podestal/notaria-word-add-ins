import axios from "axios";

/* global Word console Office */

export async function insertText(text: string) {
  // Write text to the document.
  try {
    await Word.run(async (context) => {
      let body = context.document.body;
      body.insertParagraph(text, Word.InsertLocation.end);
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export function saveDocumentFromTaskpane(): Promise<void> {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 }, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        reject(result.error);
        return;
      }

      const file = result.value;
      const sliceCount = file.sliceCount;
      const docdata: any[] = [];
      let slicesReceived = 0;

      const finalizeFailure = (error?: any) => {
        console.log("Save document error:", error);
        file.closeAsync();
        Office.context.ui.displayDialogAsync(window.location.origin + "/assets/feedback-error.html", {
          height: 30,
          width: 20,
          displayInIframe: true,
        });
        reject(error);
      };

      const getSlice = (sliceIndex: number) => {
        file.getSliceAsync(sliceIndex, (sliceResult) => {
          if (sliceResult.status !== Office.AsyncResultStatus.Succeeded) {
            finalizeFailure(sliceResult.error);
            return;
          }

          docdata[sliceIndex] = sliceResult.value.data;
          slicesReceived++;

          if (slicesReceived !== sliceCount) {
            getSlice(sliceIndex + 1);
            return;
          }

          const byteArray = docdata.reduce((acc, curr) => acc.concat(Array.from(curr)), []);
          const blob = new Blob([new Uint8Array(byteArray)], {
            type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          });

          let fileName = "document.docx";
          try {
            const documentUrl = Office.context.document.url;
            if (documentUrl) {
              const urlParts = documentUrl.split("/");
              const lastPart = urlParts[urlParts.length - 1];
              if (lastPart && lastPart.includes(".")) {
                fileName = decodeURIComponent(lastPart);
              }
            }
          } catch (error) {
            console.log("Error getting filename from URL:", error);
          }

          fileName = fileName.replace(/\(\d+\)(?=\.docx$)/i, "");
          if (!fileName.endsWith(".docx")) {
            fileName += ".docx";
          }

          const formData = new FormData();
          formData.append("file", blob, fileName);

          Office.context.ui.displayDialogAsync(
            window.location.origin + "/assets/feedback-waiting.html",
            { height: 30, width: 20, displayInIframe: true },
            (dialogResult) => {
              const waitingDialog = dialogResult.value;
              axios
                .post(`${process.env.API_BASE_URL}${process.env.API_ENDPOINT}`, formData, {
                  headers: {
                    "Content-Type": "multipart/form-data",
                    Authorization: `Token ${process.env.API_TOKEN}`,
                  },
                })
                .then(() => {
                  waitingDialog?.close();
                  Office.context.ui.displayDialogAsync(window.location.origin + "/assets/feedback-success.html", {
                    height: 30,
                    width: 20,
                    displayInIframe: true,
                  });
                  file.closeAsync();
                  resolve();
                })
                .catch((error) => {
                  waitingDialog?.close();
                  finalizeFailure(error);
                });
            }
          );
        });
      };

      getSlice(0);
    });
  });
}
