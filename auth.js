const msalConfig = {
  auth: {
    clientId: "c43fd9f3-f6a6-4b18-88e6-ee64e05db94e",
    authority: "https://login.microsoftonline.com/common",
    redirectUri: "https://aamir30091993.github.io/outlook-addin/auth.html" // Must exactly match your Azure AD app registration
  }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);
const loginRequest = {
  scopes: ["User.Read", "Mail.Read"]
};

Office.onReady(async () => {
  try {
    // Handle redirect response after login (when redirected back from Microsoft)
    const response = await msalInstance.handleRedirectPromise();

    if (response) {
      // Login success - get access token
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
      // No login response yet - trigger redirect login
      await msalInstance.loginRedirect(loginRequest);
      // After this line, the page will redirect to Microsoft login and then back to redirectUri
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
