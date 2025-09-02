// MSALWrapper.ts
import {
  PublicClientApplication,
  AuthenticationResult,
  Configuration,
  InteractionRequiredAuthError,
  AccountInfo,
} from "@azure/msal-browser";

export class MSALWrapper {
  private msalConfig: Configuration;
  private isInitialized = false;

  private msalInstance: PublicClientApplication;

  constructor(clientId: string, authority: string) {
    this.msalConfig = {
      auth: {
        clientId: clientId,
        authority: authority,
        redirectUri: `${window.location.origin}${window.location.pathname}`,
      },
      cache: {
        cacheLocation: "localStorage",
      },
    };

    this.msalInstance = new PublicClientApplication(this.msalConfig);
  }
  private async ensureInitialized(): Promise<void> {
    if (!this.isInitialized) {
      await this.msalInstance.initialize();
      this.isInitialized = true;
    }
  }
  /* Handles the logged-in user by first attempting to use SSO silent authentication.
   * If that fails, it checks the cache for the user account and attempts to acquire a token  */
  public async handleLoggedInUser(
    scopes: string[],
    userEmail: string,
  ): Promise<AuthenticationResult | undefined> {
    await this.ensureInitialized();
    try {
      // Try silent SSO to ensure MSAL has the account
      const result = await this.msalInstance.ssoSilent({
        scopes,
        loginHint: userEmail,
      });

      if (result) {
        return result;
      }
    } catch (ssoError) {
      console.log("ssoSilent failed, will try acquireTokenSilent:", ssoError);
      // No-op: proceed to check cache manually
    }

    // Try from cache next
    let userAccount: AccountInfo | null = null;
    const accounts = this.msalInstance.getAllAccounts();

    if (!accounts || accounts.length === 0) {
      console.log("No users are signed in even after ssoSilent.");
      return undefined;
    } else if (accounts.length > 1) {
      userAccount =
        accounts.find(
          (account) =>
            account.username.toLowerCase() === userEmail.toLowerCase(),
        ) ?? null;
    } else {
      userAccount = accounts[0];
    }

    if (userAccount !== null) {
      const accessTokenRequest = {
        scopes,
        account: userAccount,
      };

      return this.msalInstance
        .acquireTokenSilent(accessTokenRequest)
        .then((response) => response)
        .catch((errorinternal) => {
          console.log("acquireTokenSilent failed:", errorinternal);
          return undefined;
        });
    }

    return undefined;
  }
  /*
    public async acquireAccessToken(
        scopes: string[],
        userEmail: string,
    ): Promise<AuthenticationResult | undefined> {
        await this.ensureInitialized();
        const accessTokenRequest = {
            scopes: scopes,
            loginHint: userEmail,
        };

        return this.msalInstance
            .ssoSilent(accessTokenRequest)
            .then((response) => {
                return response;
            })
            .catch((silentError) => {
                console.log(silentError);
                if (silentError instanceof InteractionRequiredAuthError) {
                    return this.msalInstance
                        .loginPopup(accessTokenRequest)
                        .then((response) => {
                            return response;
                        })
                        .catch((error) => {
                            console.log(error);
                            return undefined;
                        });
                }
                return undefined;
            });
    }
}
*/
  /* Handles the logged-in user by first attempting to use SSO silent authentication.
  If that fails, it checks the cache for the user account and attempts to acquire a token
  */
  public async acquireAccessToken(
    scopes: string[],
    userEmail?: string,
  ): Promise<AuthenticationResult | undefined> {
    await this.ensureInitialized();

    const accounts = this.msalInstance.getAllAccounts();

    // No account? Force login
    if (!accounts || accounts.length === 0) {
      console.log("No account found. Prompting login...");
      try {
        const loginResponse = await this.msalInstance.loginPopup({
          scopes,
          loginHint: userEmail,
        });
        return loginResponse;
      } catch (loginError) {
        console.error("Login failed:", loginError);
        return undefined;
      }
    }

    // Use the first (or matching) account
    let userAccount: AccountInfo = accounts[0];
    if (userEmail && accounts.length > 1) {
      userAccount =
        accounts.find(
          (acc) => acc.username.toLowerCase() === userEmail.toLowerCase(),
        ) ?? accounts[0];
    }

    try {
      // Try silent token acquisition
      return await this.msalInstance.acquireTokenSilent({
        scopes,
        account: userAccount,
      });
    } catch (silentError) {
      console.warn("Silent token failed, trying popup:", silentError);
      if (silentError instanceof InteractionRequiredAuthError) {
        try {
          return await this.msalInstance.acquireTokenPopup({
            scopes,
            account: userAccount,
          });
        } catch (popupError) {
          console.error("Popup token request failed:", popupError);
          return undefined;
        }
      }
      return undefined;
    }
  }
}

export default MSALWrapper;
