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

function run() {
  callWebService("aamir.s@colliers.com");
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
	
	console.log(result);

    if (result) {
	
      console.log("inside if result()");	
		
      localStorage.setItem("Token", result.Token);
      localStorage.setItem("TokenID", result.TokenID);
      localStorage.setItem("UserID", result.UserID);
      localStorage.setItem("CountryID", result.CountryID);
      localStorage.setItem("Lang", result.Lang);
      localStorage.setItem("OrgSlID", result.OrgSlID);
      localStorage.setItem("Status", result.Status);
      localStorage.setItem("UserName", result.UserName);
    }
	
	console.log(localStorage);

    return "";

  } catch (err) {
    console.error("Error in callWebService:", err);
    return null;
  }
}

async function webPostMethod(postData, url) {
  console.log("Entered webPostMethod()");
  let responseText = "";

  try {
    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded"
      },
      body: postData
    });

    if (!response.ok) {
      console.error("Response status not OK:", response.status);
      return "";
    }

    responseText = await response.text();
    console.log("Response from server:", responseText);
  } catch (error) {
    console.error("Exception in webPostMethod():", error.message);
  }

  return responseText;
}