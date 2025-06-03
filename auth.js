const msalConfig = {
  auth: {
    clientId: "c43fd9f3-f6a6-4b18-88e6-ee64e05db94e",
    authority: "https://login.microsoftonline.com/common",
    redirectUri: "https://aamir30091993.github.io/outlook-addin/auth.html"
  }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);
const loginRequest = {
  scopes: ["User.Read", "Mail.Read"]
};

(async function () {
  try {
    const response = await msalInstance.loginPopup(loginRequest);
    const token = response.accessToken;
    Office.context.ui.messageParent(token);
  } catch (e) {
    console.error("Login failed:", e);
    Office.context.ui.messageParent("ERROR:" + e.message);
  }
})();
