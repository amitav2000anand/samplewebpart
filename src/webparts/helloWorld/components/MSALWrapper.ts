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

  public async handleLoggedInUser(
    scopes: string[],
    userEmail: string,
  ): Promise<AuthenticationResult | undefined> {
    let userAccount: AccountInfo | null = null;
    const accounts = this.msalInstance.getAllAccounts();

    if (accounts === null || accounts.length === 0) {
      console.log("No users are signed in");
      return undefined;
    } else if (accounts.length > 1) {
      userAccount = this.msalInstance.getAccountByUsername(userEmail);
    } else {
      userAccount = accounts[0];
    }

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

  public async acquireAccessToken(
    scopes: string[],
    userEmail: string,
  ): Promise<AuthenticationResult | undefined> {
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
