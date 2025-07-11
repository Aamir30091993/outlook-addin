

let authDialog;
let retrievedTokenID;
let userEmail;

// Storage helper keys and functions
const STORAGE_KEY = "MyAddin:SessionData";
const EXPIRY_MS = 24 * 60 * 60 * 1000; // 24 hours in ms

function loadSessionData() {
  const json = localStorage.getItem(STORAGE_KEY);
  if (!json) return { instances: {}, issued: null, tokenID: null };
  try {
    const session = JSON.parse(json);
    if (session.issued) {
      const issuedTime = new Date(session.issued).getTime();
      if (Date.now() - issuedTime > EXPIRY_MS) {
        // expired, clear storage
        localStorage.removeItem(STORAGE_KEY);
        return { instances: {}, issued: null, tokenID: null };
      }
    }
    return session;
  } catch {
    localStorage.removeItem(STORAGE_KEY);
    return { instances: {}, issued: null, tokenID: null };
  }
}

function saveSessionData(data) {
  if (!data.issued) data.issued = new Date().toISOString();
  localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
}

function makeInstanceKey(tokenID, conversationID, from, toList, emailDate) {
  // Normalize 'to' list by splitting and sorting
  const recipients = toList.split(';').map(r => r.trim().toLowerCase()).filter(r => r);
  recipients.sort();
  const normalizedTo = recipients.join(';');
  return [tokenID, conversationID, from.trim().toLowerCase(), normalizedTo].join("|");
}

function storeInstanceForConversation(tokenID, conversationID, from, to, instanceID, emailDate) {
  const session = loadSessionData();
  session.tokenID = tokenID;
  const key = makeInstanceKey(tokenID, conversationID, from, to, emailDate);
  session.instances = session.instances || {};
  session.instances[key] = { instanceID, created: new Date().toISOString() };
  saveSessionData(session);
}

function getStoredInstanceID(tokenID, conversationID, from, to, emailDate) {
  const session = loadSessionData();
  if (!session.instances) return null;
  const key = makeInstanceKey(tokenID, conversationID, from, to, emailDate);
  return session.instances[key]?.instanceID || null;
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    userEmail = Office.context.mailbox.userProfile.emailAddress;
	if (userEmail && userEmail.includes("@")) {
        userEmail = userEmail.split("@")[0];
       }
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
          await callWebService(userEmail); //userEmail //aamir.s
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
            console.log("Dialog response:", arg.message);
            authDialog.close();

            if (arg.message.startsWith("ERROR:")) {
              reject(arg.message);
            } else {
              try {
                const result = JSON.parse(arg.message);
                
                // Save accessToken & account
                localStorage.setItem("accessToken", result.accessToken);
                localStorage.setItem("msalAccount", JSON.stringify(result.account));

                // Rehydrate the MSAL instance (if available globally)
                if (window.msalInstance) {
                  window.msalInstance.setActiveAccount(result.account);
                }

                resolve(result.accessToken);
              } catch (e) {
                reject("Failed to parse auth dialog response.");
              }
            }
          });

          authDialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
            console.warn("Dialog closed or failed:", arg.error);
            reject("Dialog closed or failed.");
          });
        }
      }
    );
  });
}

async function getAccessToken() {
  try {
    const msalAccountJson = localStorage.getItem("msalAccount");
    const msalAccount = msalAccountJson ? JSON.parse(msalAccountJson) : null;

    if (!msalAccount) {
      throw new Error("No account found");
    }

    if (!window.msalInstance) {
      throw new Error("MSAL instance is not initialized");
    }

    // Set the active account if not already
    window.msalInstance.setActiveAccount(msalAccount);

    const silentRequest = {
      scopes: ["User.Read", "Mail.Read"],
      account: msalAccount
    };

    const response = await window.msalInstance.acquireTokenSilent(silentRequest);
    console.log("Silent token acquired.");
    localStorage.setItem("accessToken", response.accessToken);
    return response.accessToken;

  } catch (error) {
    console.warn("Silent token failed, using popup:", error.message);

    // Only open dialog if no tokenID already exists
    const tokenID = localStorage.getItem("TokenID");
    if (!tokenID) {
      return await loginWithDialog();
    } else {
      console.warn("TokenID already exists, skipping auth popup.");
      return localStorage.getItem("accessToken");
    }
  }
}

async function handleProceed() {
  // Hide welcome UI and show logout
  document.querySelector("header").style.display = "none";
  document.getElementById("run").style.display = "none";
  document.getElementById("logout").style.display = "block";
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";

  const item = Office.context.mailbox.item;
  
  //const convId = item.conversationId;
  
  const convId = await getConversationId(item);

  // Extract item fields
  const { mode, from, to, subject, emailDate } = await extractItemInfo();

  const tokenID = localStorage.getItem("TokenID");

  // Determine or create instanceID
  let instanceID = getStoredInstanceID(tokenID, convId, from, to, emailDate);
  if (instanceID) {
    console.log("Reusing existing instanceID:", instanceID);
  } else {
	console.log("Existing instanceID:", instanceID);
	if (instanceID == null)
	{
		instanceID = "0";
	}
  
    //const payload = { mode, from, to, subject, date, conversationId: convId, tokenID };
	//const payload = {instanceID, tokenID, from, subject, date};
	
	const payload = new URLSearchParams();   	
	//const postData = new URLSearchParams();
    payload.append("instanceID", instanceID);
    payload.append("tokenID", tokenID);
	payload.append("clientEmailAddress", from);
    payload.append("emailSubject", subject);
	payload.append("emailDate", emailDate);
	payload.append("from", from);
	payload.append("to", to);
	payload.append("conversationid", convId);
   
	
    instanceID = await callYourApi(payload);
	console.log("After api call:", instanceID);
    if (instanceID) {
      storeInstanceForConversation(tokenID, convId, from, to, instanceID, emailDate);
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

/**
 * Returns the conversationId for an item, in both read and compose modes,
 * using Microsoft Graph and MSAL.
 */
async function getConversationId(item) {
  // 1) Read mode already has it
  if (item.conversationId) {
    console.log("Using item.conversationId:", item.conversationId);
    return item.conversationId;
  }

  // 2) Compose mode: save the draft to get itemId
  const saveRes = await new Promise((resolve) => item.saveAsync(resolve));
  if (saveRes.status !== Office.AsyncResultStatus.Succeeded) {
    console.error("Draft save failed:", saveRes.error);
    return null;
  }
  const itemId = saveRes.value;
  console.log("Draft saved, ItemId:", itemId);

  // 3) Get Graph token using getAccessToken()
  let graphToken;
  try {
    graphToken = await getAccessToken(); // ← uses MSAL with fallback
  } catch (err) {
    console.error("Failed to get Graph token:", err);
    return null;
  }

  // 4) Call Microsoft Graph to get conversationId
  const graphUrl = `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(itemId)}?$select=conversationId`;

  try {
    const response = await fetch(graphUrl, {
      headers: {
        Authorization: `Bearer ${graphToken}`,
        Accept: "application/json"
      }
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error("Graph API failed:", response.status, errorText);
      return null;
    }

    const data = await response.json();
    console.log("Fetched conversationId from Graph:", data.conversationId);
    return data.conversationId;

  } catch (fetchError) {
    console.error("Error calling Graph API:", fetchError);
    return null;
  }
}


async function extractItemInfo() {
  const item = Office.context.mailbox.item;
  const isCompose = !!item.subject.getAsync;
  let mode, from, to, subject, emailDate;

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
    //date = new Date().toISOString();
	emailDate = formatEmailDate(new Date().toISOString());
  } else {
    mode = "Read";
    from = item.from?.emailAddress || "";
    to = (item.to || []).map((x) => x.emailAddress).join("; ");
    subject = item.subject || "";
    //date = item.dateTimeCreated
      //? item.dateTimeCreated.toISOString()
      //: new Date().toISOString();
	emailDate = formatEmailDate(item.dateTimeCreated.toISOString());
	
	console.log(emailDate);
  }
  return { mode, from, to, subject, emailDate };
}

function formatEmailDate(isoString) {
  const d = new Date(isoString);

  // helper to zero-pad numbers under 10
  const pad = (n) => n.toString().padStart(2, "0");

  const day     = pad(d.getDate());
  const month   = pad(d.getMonth() + 1);
  const year    = d.getFullYear();

  const hours   = pad(d.getHours());
  const minutes = pad(d.getMinutes());
  const seconds = pad(d.getSeconds());

  return `${day}/${month}/${year} ${hours}:${minutes}:${seconds}`;
}

async function callYourApi(data) {
  try {
    console.log("Calling API with data:", data);
    const response = await fetch(
      "https://uat-uae-ezconnect.colliersasia.com/Instance/insertUpdateInstance",
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: data.toString()
      }
    );
    if (!response.ok) throw new Error(response.statusText);
    const json = await response.json();
    console.log("API response:", json);
    // Store responseMsg as instanceID
    return json.responseMsg;
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
