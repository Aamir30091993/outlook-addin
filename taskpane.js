/* global Office, document */

let authDialog;
let retrievedTokenID;

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = async () => {
      try {
        const token = await loginWithDialog();
        console.log("Token received in taskpane.js:", token);
        await run(); // ✅ Await run()
	 
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
    await callWebService("aamir.s@colliers.com"); // ✅ Wait until token is set

	document.querySelector("header").style.display = "none";
	document.getElementById("run").style.display = "none";
	document.getElementById("sideload-msg").style.display = "none";

    // ✅ Set after the token is stored
	const iframe = document.getElementById("webFrame");
	iframe.style.display = "block";
	
	 // ✅ Now retrievedTokenID is set correctly
    retrievedTokenID = localStorage.getItem("TokenID");
    console.log("Setting iframe with tokenID:", retrievedTokenID);
	
	console.log(" EncodedURIComponent retrievedTokenID");
	console.log(encodeURIComponent(retrievedTokenID));

	//calling the webpage part	   	 
	iframe.src = "https://uat-uae-ezconnect.colliersasia.com/?tokenID=" + encodeURIComponent(retrievedTokenID) + "&instanceID=0";
  
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
	  
      // Setting a value in localStorage	  
	   	
      localStorage.setItem("Token", result.token);
      localStorage.setItem("TokenID", result.tokenid);
      localStorage.setItem("UserID", result.userID);
      localStorage.setItem("CountryID", result.countryID);
      localStorage.setItem("Lang", result.lang);
      localStorage.setItem("OrgSlID", result.orgSlID);
      localStorage.setItem("Status", result.status);
      localStorage.setItem("UserName", result.userName);
    }
	
	// Getting the value from localStorage
       const retrievedToken = localStorage.getItem("Token");	
       console.log(retrievedToken);
	   
	   retrievedTokenID = localStorage.getItem("TokenID");	
       console.log(retrievedTokenID);
	   
	   const retrievedUserID = localStorage.getItem("UserID");	
       console.log(retrievedUserID);
	   
	   const retrievedCountryID = localStorage.getItem("CountryID");	
       console.log(retrievedCountryID);
	   
	   const retrievedLang = localStorage.getItem("Lang");	
       console.log(retrievedLang);
	   
	   const retrievedOrgSlID = localStorage.getItem("OrgSlID");	
       console.log(retrievedOrgSlID);
	   
	   const retrievedStatus = localStorage.getItem("Status");	
       console.log(retrievedStatus);
	   
	   const retrievedUserName = localStorage.getItem("UserName");	
       console.log(retrievedUserName);
	   
	 
	   
	   

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