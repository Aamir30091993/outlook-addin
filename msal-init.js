// msal-init.js (loaded only in taskpane.html)

const msalConfig = {
  auth: {
    clientId: "c43fd9f3-f6a6-4b18-88e6-ee64e05db94e",
    authority: "https://login.microsoftonline.com/common",
    redirectUri: "https://aamir30091993.github.io/outlook-addin/auth.html"
  }
};
window.msalInstance = new msal.PublicClientApplication(msalConfig);
