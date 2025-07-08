const msalConfig = {
  auth: {
    clientId: "c43fd9f3-f6a6-4b18-88e6-ee64e05db94e",
    authority: "https://login.microsoftonline.com/common",
    redirectUri: "https://aamir30091993.github.io/outlook-addin/auth.html"
  }
};

window.msalInstance = new msal.PublicClientApplication(msalConfig);
const loginRequest = {
  scopes: ["User.Read", "Mail.ReadWrite"]
};

// Immediately run loginPopup in this dialog
(async () => {
  try {
    const response = await window.msalInstance.loginPopup(loginRequest);
    Office.context.ui.messageParent(response.accessToken);
  } catch (e) {
    Office.context.ui.messageParent("ERROR:" + e.message);
  }
})();
