/* global Office, document */

let authDialog;
let retrievedTokenID;
let userEmail;

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
