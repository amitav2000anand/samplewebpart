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
  /*
    public async handleLoggedInUser(
      scopes: string[],
      userEmail: string,
    ):
      Promise<AuthenticationResult | undefined> {
      let userAccount: AccountInfo | null = null;
      const accounts = this.msalInstance.getAllAccounts();
  
      if (accounts === null || accounts.length === 0 || userEmail === null) {
        console.log("No users are signed in");
        return undefined;
      } else if (accounts.length > 1) {
        userAccount = accounts.find(
          (account) => account.username.toLowerCase() === userEmail.toLowerCase()
        ) ?? null;
      } else {
        userAccount = accounts[0];
      }//{
        //userAccount = this.msalInstance.getAccountByUsername(userEmail);
      //} else {
       // userAccount = accounts[0];
      //}
  
      if (userAccount !== null) {
        const accessTokenRequest = {
          scopes: scopes,
          account: userAccount,
        };
  
        return this.msalInstance
          .acquireTokenSilent(accessTokenRequest)
          .then((response) => {
            return response;
          })
          .catch((errorinternal) => {
            console.log(errorinternal);
            return undefined;
          });
      }
      return undefined;
    }
  */
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
      userAccount = accounts.find(
        (account) => account.username.toLowerCase() === userEmail.toLowerCase()
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

export default MSALWrapper;
