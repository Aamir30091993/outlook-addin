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

Office.onReady(async () => {
  try {
    // Handle redirect response after login
    const response = await msalInstance.handleRedirectPromise();

    if (response) {
      // We got a response after redirect login
      const token = response.accessToken;

      if (
        Office &&
        Office.context &&
        Office.context.ui &&
        typeof Office.context.ui.messageParent === "function"
      ) {
        Office.context.ui.messageParent(token);
      } else {
        console.warn("Office.context.ui.messageParent is not available.");
      }
    } else {
      // No response yet — trigger redirect login
      await msalInstance.loginRedirect(loginRequest);
      // The page will redirect, so code after this usually won't run
    }
  } catch (e) {
    console.error("Login failed:", e);
    if (
      Office &&
      Office.context &&
      Office.context.ui &&
      typeof Office.context.ui.messageParent === "function"
    ) {
      Office.context.ui.messageParent("ERROR:" + e.message);
    }
  }
});
