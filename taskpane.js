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
  
  //const iframe = document.getElementById("webFrame");
  //iframe.src = "https://ezconnect.colliersasia.com/login.aspx"; // Replace with your page URL
  
  callWebService ("aamir.s@benchmarksolution.com")
}

async function callWebService(username) {
  try {
    console.log("Entered callWebService");

    const postData = new URLSearchParams();
    postData.append("UserName", username);
    postData.append("Password", "123"); // Hardcoded as in original code

    // const response = await fetch("https://uat-uae-ezconnect.colliersasia.com/api/Token?", {
      // method: "POST",
      // headers: {
        // "Content-Type": "application/x-www-form-urlencoded"
      // },
      // body: postData
    // });
	
	const URL = "https://uat-uae-ezconnect.colliersasia.com/api/Token?";

    // const resultText = await response.text();
    // if (!resultText) {
      // console.warn("Empty response from API.");
      // return null;
    // }

    //Convert JSON to XML if needed, or parse as JSON directly:
    // const result = JSON.parse(resultText);

    // console.log("Received token data:", result);
	
	const json = await webPostMethod(postData, URL);
	
	const result = JSON.parse(json);

	 if (result)
	 {
		// Store token data in localStorage or a secure place
		localStorage.setItem("Token", result.Token);
		localStorage.setItem("TokenID", result.TokenID);
		localStorage.setItem("UserID", result.UserID);
		localStorage.setItem("CountryID", result.CountryID);
		localStorage.setItem("Lang", result.Lang);
		localStorage.setItem("OrgSlID", result.OrgSlID);
		localStorage.setItem("Status", result.Status);
		localStorage.setItem("UserName", result.UserName);
	}

    // Example of showing login success in task pane
    //document.getElementById("loginStatus").innerText = "Login Successful";
    return "";

  } catch (err) {
    console.error("Error in callWebService:", err);

    // Display fallback UI in taskpane
   // document.getElementById("errorDisplay").innerText = "Login failed or network error.";
    return null;
  }
}

async function webPostMethod(postData, url) {
  console.log("Entered webPostMethod()");
  let responseText = "";

  try {
    // Set a timeout (like in .Timeout = 50000 ms in C#)
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 50000);

    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        "User-Agent": "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 7.1; Trident/5.0)",
        "Accept": "*/*"
      },
      body: new URLSearchParams(postData),
      signal: controller.signal
    });

    clearTimeout(timeoutId);

    if (!response.ok) {
      console.error("Response status not OK:", response.status);
      //displayNetworkErrorUI(); // See below
      return "";
    }

    responseText = await response.text();
    console.log("Response from server:", responseText);

  } catch (error) {
    console.error("Exception in webPostMethod():", error.message);
    //displayNetworkErrorUI();
  }

  return responseText;
}