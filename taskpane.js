/* global Office, document */

let authDialog;
let retrievedTokenID;
let userEmail;

// Storage helper keys and functions
const STORAGE_KEY = "MyAddin:SessionData";

function loadSessionData() {
  const jsonSK = localStorage.getItem(STORAGE_KEY);
  console.log("jsonSK");
  console.log(jsonSK);
  return jsonSK ? JSON.parse(jsonSK) : null;
}

function saveSessionData(data) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
  console.log("saveSessionData");
  console.log(data);
}

function storeInstanceForConversation(tokenID, conversationID, instanceID) {
  let session = loadSessionData();
  if (!session || session.tokenID !== tokenID) {
    session = {
      tokenID: tokenID,
      issued: new Date().toISOString(),
      conversations: {}
    };
  }
  session.conversations[conversationID] = {
    instanceID: instanceID,
    created: new Date().toISOString()
  };
  saveSessionData(session);
}

function getStoredInstanceID(conversationID) {
  const session = loadSessionData();
  if (session && session.conversations[conversationID]) {
    return session.conversations[conversationID].instanceID;
  }
  return null;
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    userEmail = Office.context.mailbox.userProfile.emailAddress;
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // Logout handler
    document.getElementById("logout").onclick = () => {
      console.log("Logging out...");
      localStorage.clear();
      document.querySelector("header").style.display = "block";
      document.getElementById("run").style.display = "block";
      document.getElementById("logout").style.display = "none";
      document.getElementById("webFrame").style.display = "none";
      document.getElementById("webFrame").src = "";
    };

    // Handle proceed click
    document.getElementById("run").onclick = async () => {
      try {
        await handleProceed();
      } catch (e) {
        console.error("Error on proceed:", e);
      }
    };
  }
});

async function handleProceed() {
  // Hide welcome UI and show logout
  document.querySelector("header").style.display = "none";
  document.getElementById("run").style.display = "none";
  document.getElementById("logout").style.display = "block";
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";

  const item = Office.context.mailbox.item;
  const convId = item.conversationId;

  // Extract item fields
  const { mode, from, to, subject, date } = await extractItemInfo();

  // Determine or create instanceID
  let instanceID = getStoredInstanceID(convId);
  if (instanceID) {
    console.log("Reusing existing instanceID:", instanceID);
  } else {
    const tokenID = localStorage.getItem("TokenID");
    const payload = { mode, from, to, subject, date, conversationId: convId, tokenID };
    instanceID = await callYourApi(payload);
    if (instanceID) {
      storeInstanceForConversation(tokenID, convId, instanceID);
      console.log("Stored new instanceID for conversation", convId);
    }
  }


  // Load iframe
  const retrievedTokenID = localStorage.getItem("TokenID");
  const iframe = document.getElementById("webFrame");
  iframe.style.display = "block";
  iframe.src =
    `https://uat-uae-ezconnect.colliersasia.com/?tokenID=${encodeURIComponent(
      retrievedTokenID
    )}&instanceID=${encodeURIComponent(instanceID)}`;
}

async function extractItemInfo() {
  const item = Office.context.mailbox.item;
  const isCompose = !!item.subject.getAsync;
  let mode, from, to, subject, date;

  if (isCompose) {
    mode = "Compose";
    from = userEmail;
    to = await new Promise((r) =>
      item.to.getAsync((res) =>
        r(
          res.status === Office.AsyncResultStatus.Succeeded
            ? res.value.map((x) => x.emailAddress).join("; ")
            : ""
        )
      )
    );
    subject = await new Promise((r) =>
      item.subject.getAsync((res) => r(res.status === Office.AsyncResultStatus.Succeeded ? res.value : ""))
    );
    date = new Date().toISOString();
  } else {
    mode = "Read";
    from = item.from?.emailAddress || "";
    to = (item.to || []).map((x) => x.emailAddress).join("; ");
    subject = item.subject || "";
    date = item.dateTimeCreated
      ? item.dateTimeCreated.toISOString()
      : new Date().toISOString();
  }
  return { mode, from, to, subject, date };
}

async function callYourApi(data) {
  try {
    console.log("Calling API with data:", data);
    const response = await fetch(
      "https://uat-uae-ezconnect.colliersasia.com/Instance/insertUpdateInstance",
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(data)
      }
    );
    if (!response.ok) throw new Error(response.statusText);
    const json = await response.json();
	console.log("InstanceID :", json);
    return json.uniqueID;
  } catch (e) {
    console.error("API error:", e);
    return null;
  }
}
