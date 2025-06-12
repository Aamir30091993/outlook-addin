/* global Office, document */

let authDialog;

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = async () => {
      try {
        const token = await loginWithDialog();
        console.log("Token received in taskpane.js:", token);
        run();
      } catch (error) {
        console.error("Login failed or dialog error:", error);
      }
    };
  }
});

function loginWithDialog() {
  return new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(
      "https://aamir30091993.github.io/outlook-addin/auth.html",
      { height: 60, width: 30 },  // Removed displayInIframe option here
      function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          reject(asyncResult.error.message);
        } else {
          authDialog = asyncResult.value;
          authDialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
            console.log("Token received:", arg.message);
            authDialog.close();
            resolve(arg.message); // this is your access token or error message
          });
          authDialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
            console.warn("Dialog event received:", arg.error);
          });
        }
      }
    );
  });
}

function run() {
/*   const item = Office.context.mailbox.item;
  const insertAt = document.getElementById("item-subject");
  insertAt.innerHTML = "";

  const label = document.createElement("b");
  label.textContent = "Subject: ";
  insertAt.appendChild(label);
  insertAt.appendChild(document.createElement("br"));
  insertAt.appendChild(document.createTextNode(item.subject || "No subject"));
  insertAt.appendChild(document.createElement("br"));
  insertAt.appendChild(document.createTextNode("Item ID: " + (item.itemId || "N/A")));
  insertAt.appendChild(document.createElement("br"));

  if (item.dateTimeCreated) {
    insertAt.appendChild(document.createTextNode("Created: " + item.dateTimeCreated.toString()));
    insertAt.appendChild(document.createElement("br"));
  }

  insertAt.appendChild(document.createTextNode("Item Type: " + (item.itemType || "Unknown")));
  insertAt.appendChild(document.createElement("br"));

  if (item.from && item.from.emailAddress) {
    insertAt.appendChild(document.createTextNode("From: " + item.from.emailAddress));
    insertAt.appendChild(document.createElement("br"));
  } */
  
  const iframe = document.getElementById("webFrame");
  iframe.src = "https://www.google.com/"; // Replace with your page URL
}
