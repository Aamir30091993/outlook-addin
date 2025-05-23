/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

  async function run() {
  const item = Office.context.mailbox.item;
  const insertAt = document.getElementById("item-subject");

  // Clear existing content
  insertAt.innerHTML = "";

  // Add Subject
  const label = document.createElement("b");
  label.textContent = "Subject: ";
  insertAt.appendChild(label);
  insertAt.appendChild(document.createElement("br"));
  insertAt.appendChild(document.createTextNode(item.subject || "No subject"));
  insertAt.appendChild(document.createElement("br"));

  // Add itemId
  insertAt.appendChild(document.createTextNode("Item ID: " + (item.itemId || "N/A")));
  insertAt.appendChild(document.createElement("br"));

  // Add dateTimeCreated (only available in Read mode)
  if (item.dateTimeCreated) {
    insertAt.appendChild(document.createTextNode("Created: " + item.dateTimeCreated.toString()));
    insertAt.appendChild(document.createElement("br"));
  }

  // Add itemType
  insertAt.appendChild(document.createTextNode("Item Type: " + (item.itemType || "Unknown")));
  insertAt.appendChild(document.createElement("br"));

  // Add body (only available in Compose mode)
  if (item.body && item.body.getAsync) {
    item.body.getAsync(Office.CoercionType.Text, result => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        insertAt.appendChild(document.createTextNode("Body (Compose mode): " + result.value));
        insertAt.appendChild(document.createElement("br"));
      } else {
        insertAt.appendChild(document.createTextNode("Error reading body: " + result.error.message));
        insertAt.appendChild(document.createElement("br"));
      }
    });
  }

  // Add from (only available in Read mode)
  if (item.from && item.from.emailAddress) {
    insertAt.appendChild(document.createTextNode("From: " + item.from.emailAddress));
    insertAt.appendChild(document.createElement("br"));
  }
}
