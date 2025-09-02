import { AuthenticationResult } from "@azure/msal-browser";
export declare class MSALWrapper {
    private msalConfig;
    private isInitialized;
    private msalInstance;
    constructor(clientId: string, authority: string);
    private ensureInitialized;
    handleLoggedInUser(scopes: string[], userEmail: string): Promise<AuthenticationResult | undefined>;
    acquireAccessToken(scopes: string[], userEmail?: string): Promise<AuthenticationResult | undefined>;
}
export default MSALWrapper;
//# sourceMappingURL=MSALWrapper.d.ts.map