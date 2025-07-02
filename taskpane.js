/* global Office, document */

let authDialog;
let retrievedTokenID;
let userEmail;

// Storage helper keys and functions
const STORAGE_KEY = "MyAddin:SessionData";

  function loadSessionData() {
   const jsonSK = localStorage.getItem(STORAGE_KEY);
   return jsonSK ? JSON.parse(jsonSK) : null;
   
   console.log("jsonSK:", jsonSK);
}

  function saveSessionData(data) {
   localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
   
   console.log("SaveSessionData:", data);
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
	  
	  //Logout
	  document.getElementById("logout").onclick = () => {
	  console.log("Logging out...");
	  localStorage.clear();

	  // Show welcome UI again
	  document.querySelector("header").style.display = "block";
	  document.getElementById("run").style.display = "block";
	  document.getElementById("logout").style.display = "none";
	  document.getElementById("webFrame").style.display = "none";
	  document.getElementById("webFrame").src = "";
	};
	//Logout
	
	   // Every time the user opens a message (or switches messages),
    // this event fires and we can extract the fields.
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.ItemChanged,
      () => extractItemInfo()
    );

    // Also run once on initial load
    extractItemInfo();
	  
    userEmail = Office.context.mailbox.userProfile.emailAddress;
    console.log("userEmail:", userEmail);

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    document.getElementById("run").onclick = async () => {
      try {
        const existingTokenID = localStorage.getItem("TokenID");

        if (existingTokenID) {
          console.log("TokenID found in localStorage. Skipping Azure login.");
		  console.log(existingTokenID);
          runWithToken(existingTokenID);
        } else {
          const token = await loginWithDialog();
          console.log("Token received in taskpane.js:", token);
          await run();
        }
      } catch (error) {
        console.error("Login failed or dialog error:", error);
      }
    };
  }
});

async function extractItemInfo() {
  const item = Office.context.mailbox.item;
  let mode, from, to, subject, date;

  // Compose mode has async getters
  const isCompose = !!item.subject.getAsync;

  if (isCompose) {
    mode = "Compose";
    from = Office.context.mailbox.userProfile.emailAddress;

    to = await new Promise(resolve =>
      item.to.getAsync(r =>
        resolve(
          r.status === Office.AsyncResultStatus.Succeeded
            ? r.value.map(x => x.emailAddress).join("; ")
            : "<unable to read Recipients>"
        )
      )
    );

    subject = await new Promise(resolve =>
      item.subject.getAsync(r =>
        resolve(
          r.status === Office.AsyncResultStatus.Succeeded
            ? r.value
            : "<unable to read Subject>"
        )
      )
    );

    date = new Date().toLocaleString();
  } else {
    mode = "Read";
    from = item.from?.emailAddress || "<no from>";
    to = (item.to || []).map(x => x.emailAddress).join("; ");
    subject = item.subject || "";
    date = item.dateTimeCreated
      ? item.dateTimeCreated.toLocaleString()
      : "<no date>";
  }
  
  console.log("Mode:", mode);
  console.log("From:", from);
  console.log("To:", to);
  console.log("Subject:", subject);
  console.log("Date:", date);
  
  
  //InstanceIDmapping part  - API call and storing the instance ID
  //Grab the conversationId
  const convId = item.conversationId;
  console.log("Conversation ID:", convId);
  
  // Retrieve or generate instanceID
  let instanceID = getStoredInstanceID(convId);
  if (instanceID) {
    console.log("Reusing existing instanceID:", instanceID);
  }
  else {
    //const payload = { mode, from, to, subject, date, conversationId: convId };
	
	const payload = new URLSearchParams();
    payload.append("instanceID", instanceID);
    payload.append("tokenID", localStorage.getItem("TokenID"));
	payload.append("clientEmailAddress", from);
	payload.append("emailSubject", localStorage.getItem("subject"));
	payload.append("emailDate", localStorage.getItem("date"));
	
    instanceID = await callYourApi(payload);
	
	console.log("ExtractedInstanceID from API:", instanceID)
    if (instanceID) {
      storeInstanceForConversation(
        localStorage.getItem("TokenID"),
        token,
        convId,
        instanceID
      );
      console.log("Stored new instanceID for conversation", convId);
    }
  }
  //InstanceIDmapping part  - API call and storing the instance ID
}


async function callYourApi(data) {
  try {
    const response = await fetch("https://uat-uae-ezconnect.colliersasia.com/Instance/insertUpdateInstance?", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(data)
    });
    if (!response.ok) throw new Error(response.statusText);
    const json = await response.json();
    return json.uniqueId;
  } catch (e) {
    console.error("API error:", e);
    return null;
  }
}





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

async function run() {
  await callWebService("aamir.s"); //userEmail // ✅ Replaced hardcoded email

  const iframe = document.getElementById("webFrame");

  document.querySelector("header").style.display = "none";
  document.getElementById("run").style.display = "none";
  document.getElementById("logout").style.display = "block"; // ✅
  document.getElementById("sideload-msg").style.display = "none";
  iframe.style.display = "block";	

  retrievedTokenID = localStorage.getItem("TokenID");
  console.log("Setting iframe with tokenID:", retrievedTokenID);

 iframe.src = "https://uat-uae-ezconnect.colliersasia.com/?tokenID=" + encodeURIComponent(retrievedTokenID) + "&instanceID=0";
}

// 🔁 Reused when token already exists
function runWithToken(tokenID) {
  document.querySelector("header").style.display = "none";
  document.getElementById("run").style.display = "none";
   document.getElementById("logout").style.display = "block"; // ✅ Fix added
  document.getElementById("sideload-msg").style.display = "none";

  const iframe = document.getElementById("webFrame");
  iframe.style.display = "block";

 iframe.src = "https://uat-uae-ezconnect.colliersasia.com/?tokenID=" + encodeURIComponent(tokenID) + "&instanceID=0";
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
