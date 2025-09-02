var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
// MSALWrapper.ts
import { PublicClientApplication, InteractionRequiredAuthError, } from "@azure/msal-browser";
var MSALWrapper = /** @class */ (function () {
    function MSALWrapper(clientId, authority) {
        this.isInitialized = false;
        this.msalConfig = {
            auth: {
                clientId: clientId,
                authority: authority,
                redirectUri: "".concat(window.location.origin).concat(window.location.pathname),
            },
            cache: {
                cacheLocation: "localStorage",
            },
        };
        this.msalInstance = new PublicClientApplication(this.msalConfig);
    }
    MSALWrapper.prototype.ensureInitialized = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!!this.isInitialized) return [3 /*break*/, 2];
                        return [4 /*yield*/, this.msalInstance.initialize()];
                    case 1:
                        _a.sent();
                        this.isInitialized = true;
                        _a.label = 2;
                    case 2: return [2 /*return*/];
                }
            });
        });
    };
    /* Handles the logged-in user by first attempting to use SSO silent authentication.
     * If that fails, it checks the cache for the user account and attempts to acquire a token  */
    MSALWrapper.prototype.handleLoggedInUser = function (scopes, userEmail) {
        var _a;
        return __awaiter(this, void 0, void 0, function () {
            var result, ssoError_1, userAccount, accounts, accessTokenRequest;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0: return [4 /*yield*/, this.ensureInitialized()];
                    case 1:
                        _b.sent();
                        _b.label = 2;
                    case 2:
                        _b.trys.push([2, 4, , 5]);
                        return [4 /*yield*/, this.msalInstance.ssoSilent({
                                scopes: scopes,
                                loginHint: userEmail,
                            })];
                    case 3:
                        result = _b.sent();
                        if (result) {
                            return [2 /*return*/, result];
                        }
                        return [3 /*break*/, 5];
                    case 4:
                        ssoError_1 = _b.sent();
                        console.log("ssoSilent failed, will try acquireTokenSilent:", ssoError_1);
                        return [3 /*break*/, 5];
                    case 5:
                        userAccount = null;
                        accounts = this.msalInstance.getAllAccounts();
                        if (!accounts || accounts.length === 0) {
                            console.log("No users are signed in even after ssoSilent.");
                            return [2 /*return*/, undefined];
                        }
                        else if (accounts.length > 1) {
                            userAccount =
                                (_a = accounts.find(function (account) {
                                    return account.username.toLowerCase() === userEmail.toLowerCase();
                                })) !== null && _a !== void 0 ? _a : null;
                        }
                        else {
                            userAccount = accounts[0];
                        }
                        if (userAccount !== null) {
                            accessTokenRequest = {
                                scopes: scopes,
                                account: userAccount,
                            };
                            return [2 /*return*/, this.msalInstance
                                    .acquireTokenSilent(accessTokenRequest)
                                    .then(function (response) { return response; })
                                    .catch(function (errorinternal) {
                                    console.log("acquireTokenSilent failed:", errorinternal);
                                    return undefined;
                                })];
                        }
                        return [2 /*return*/, undefined];
                }
            });
        });
    };
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
    MSALWrapper.prototype.acquireAccessToken = function (scopes, userEmail) {
        var _a;
        return __awaiter(this, void 0, void 0, function () {
            var accounts, loginResponse, loginError_1, userAccount, silentError_1, popupError_1;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0: return [4 /*yield*/, this.ensureInitialized()];
                    case 1:
                        _b.sent();
                        accounts = this.msalInstance.getAllAccounts();
                        if (!(!accounts || accounts.length === 0)) return [3 /*break*/, 5];
                        console.log("No account found. Prompting login...");
                        _b.label = 2;
                    case 2:
                        _b.trys.push([2, 4, , 5]);
                        return [4 /*yield*/, this.msalInstance.loginPopup({
                                scopes: scopes,
                                loginHint: userEmail,
                            })];
                    case 3:
                        loginResponse = _b.sent();
                        return [2 /*return*/, loginResponse];
                    case 4:
                        loginError_1 = _b.sent();
                        console.error("Login failed:", loginError_1);
                        return [2 /*return*/, undefined];
                    case 5:
                        userAccount = accounts[0];
                        if (userEmail && accounts.length > 1) {
                            userAccount =
                                (_a = accounts.find(function (acc) { return acc.username.toLowerCase() === userEmail.toLowerCase(); })) !== null && _a !== void 0 ? _a : accounts[0];
                        }
                        _b.label = 6;
                    case 6:
                        _b.trys.push([6, 8, , 13]);
                        return [4 /*yield*/, this.msalInstance.acquireTokenSilent({
                                scopes: scopes,
                                account: userAccount,
                            })];
                    case 7: 
                    // Try silent token acquisition
                    return [2 /*return*/, _b.sent()];
                    case 8:
                        silentError_1 = _b.sent();
                        console.warn("Silent token failed, trying popup:", silentError_1);
                        if (!(silentError_1 instanceof InteractionRequiredAuthError)) return [3 /*break*/, 12];
                        _b.label = 9;
                    case 9:
                        _b.trys.push([9, 11, , 12]);
                        return [4 /*yield*/, this.msalInstance.acquireTokenPopup({
                                scopes: scopes,
                                account: userAccount,
                            })];
                    case 10: return [2 /*return*/, _b.sent()];
                    case 11:
                        popupError_1 = _b.sent();
                        console.error("Popup token request failed:", popupError_1);
                        return [2 /*return*/, undefined];
                    case 12: return [2 /*return*/, undefined];
                    case 13: return [2 /*return*/];
                }
            });
        });
    };
    return MSALWrapper;
}());
export { MSALWrapper };
export default MSALWrapper;
//# sourceMappingURL=MSALWrapper.js.map