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

Office.onReady(async () => {
  try {
    const response = await msalInstance.loginPopup(loginRequest);
    const token = response.accessToken;

    if (
      Office?.context?.ui?.messageParent
    ) {
      Office.context.ui.messageParent(token);
    } else {
      console.warn("messageParent is not available.");
    }

  } catch (e) {
    console.error("Login failed:", e);

    if (
      Office?.context?.ui?.messageParent
    ) {
      Office.context.ui.messageParent("ERROR:" + e.message);
    }
  }
});
