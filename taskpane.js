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
        const tokenID = localStorage.getItem("TokenID");
        if (!tokenID) {
          // Prompt Azure login and get token
          const token = await loginWithDialog();
          console.log("Token received:", token);
          // Store token in localStorage via callWebService
          await callWebService("aamir.s"); //userEmail
        }
        // Now tokenID exists
        await handleProceed();
      } catch (e) {
        console.error("Error during authentication or proceed:", e);
      }
    };
  }
});

function loginWithDialog() {
  return new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(
      "https://aamir30091993.github.io/outlook-addin/auth.html",
      { height: 60, width: 30 },
      function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          reject(asyncResult.error.message);
        } else {
          authDialog = asyncResult.value;
          authDialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
            console.log("Token received:", arg.message);
            authDialog.close();
            resolve(arg.message);
          });
          authDialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
            console.warn("Dialog event received:", arg.error);
          });
        }
      }
    );
  });
}

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
	console.log("Existing instanceID:", instanceID);
	if (instanceID == null)
	{
		instanceID = "0";
	}
    const tokenID = localStorage.getItem("TokenID");
    //const payload = { mode, from, to, subject, date, conversationId: convId, tokenID };
	//const payload = {instanceID, tokenID, from, subject, date};
	
	const payload = new URLSearchParams();   	
	//const postData = new URLSearchParams();
    payload.append("instanceID", instanceID);
    payload.append("tokenID", tokenID);
	payload.append("clientEmailAddress", from);
    payload.append("emailSubject", subject);
	payload.append("emailDate", date);
	payload.append("from", from);
	payload.append("to", to);
	payload.append("conversationid", convId);
   
	
    instanceID = await callYourApi(payload);
	console.log("After api call:", instanceID);
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
        headers: {
          "Content-Type": "application/x-www-form-urlencoded"
        },
        body: data.toString()  // URLSearchParams or similar
      }
    );
    if (!response.ok) throw new Error(response.statusText);
    const json = await response.json();
    console.log("InstanceID response:", json);
    return json.uniqueID;    // match the exact property name from your API
  } catch (e) {
    console.error("API error:", e);
    return null;
  }
}

async function callWebService(username) {
  try {
    console.log("Entered callWebService");

    const postData = new URLSearchParams();
    postData.append("UserName", username);
    postData.append("Password", "123");

    const url = "https://uat-uae-ezconnect.colliersasia.com/api/Token?";

    const json = await webPostMethod(postData, url);
    console.log(json);

    if (!json || json.trim() === "") {
      console.warn("Empty or invalid response from API.");
      return null;
    }

    let result;
    try {
      result = JSON.parse(json);
    } catch (e) {
      console.error("Invalid JSON format:", json);
      return null;
    }

    if (result) {
      localStorage.setItem("Token", result.token);
      localStorage.setItem("TokenID", result.tokenid);
      localStorage.setItem("UserID", result.userID);
      localStorage.setItem("CountryID", result.countryID);
      localStorage.setItem("Lang", result.lang);
      localStorage.setItem("OrgSlID", result.orgSlID);
      localStorage.setItem("Status", result.status);
      localStorage.setItem("UserName", result.userName);
    }

    return "";
  } catch (err) {
    console.error("Error in callWebService:", err);
    return null;
  }
}

async function webPostMethod(postData, url) {
  console.log("Entered webPostMethod()");
  try {
    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body: postData,
    });

    if (!response.ok) {
      console.error("Response status not OK:", response.status);
      return "";
    }

    const responseText = await response.text();
    console.log("Response from server:", responseText);
    return responseText;
  } catch (error) {
    console.error("Exception in webPostMethod():", error.message);
    return "";
  }
}
