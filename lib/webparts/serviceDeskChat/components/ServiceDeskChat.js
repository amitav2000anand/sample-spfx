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
import * as React from "react";
import { useState } from "react";
import { createDirectLine, renderWebChat, createStore, } from "botframework-webchat";
import { Spinner } from "office-ui-fabric-react/lib/Spinner";
import { useRef, useEffect } from "react";
import MSALWrapper from "../../../utils/MSALWrapper";
import styles from "./ServiceDeskChat.module.scss";
var ServiceDeskChat = function (props) {
    var webChatRef = useRef(null);
    var loadingSpinnerRef = useRef(null);
    // ✅ Store directline so we can send Start Over event
    var _a = useState(null), directlineInstance = _a[0], setDirectlineInstance = _a[1];
    var botURL = props.botURL;
    var environmentEndPoint = botURL.slice(0, botURL.indexOf("/powervirtualagents"));
    var apiVersion = botURL.slice(botURL.indexOf("api-version")).split("=")[1];
    var regionalChannelSettingsURL = "".concat(environmentEndPoint, "/powervirtualagents/regionalchannelsettings?api-version=").concat(apiVersion);
    var getOAuthCardResourceUri = function (activity) {
        var _a;
        var attachment = (_a = activity === null || activity === void 0 ? void 0 : activity.attachments) === null || _a === void 0 ? void 0 : _a[0];
        if ((attachment === null || attachment === void 0 ? void 0 : attachment.contentType) === "application/vnd.microsoft.card.oauth" &&
            attachment.content.tokenExchangeResource) {
            return attachment.content.tokenExchangeResource.uri;
        }
    };
    useEffect(function () {
        var renderBot = function () { return __awaiter(void 0, void 0, void 0, function () {
            var MSALWrapperInstance, responseToken, token, regionalChannelURL, regionalResponse, data, directline, response, conversationInfo, store, canvasStyleOptions;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        MSALWrapperInstance = new MSALWrapper(props.clientID, props.authority);
                        return [4 /*yield*/, MSALWrapperInstance.handleLoggedInUser([props.customScope], props.userEmail)];
                    case 1:
                        responseToken = _a.sent();
                        if (!!responseToken) return [3 /*break*/, 3];
                        return [4 /*yield*/, MSALWrapperInstance.acquireAccessToken([props.customScope], props.userEmail)];
                    case 2:
                        responseToken = _a.sent();
                        _a.label = 3;
                    case 3:
                        token = (responseToken === null || responseToken === void 0 ? void 0 : responseToken.accessToken) || null;
                        return [4 /*yield*/, fetch(regionalChannelSettingsURL)];
                    case 4:
                        regionalResponse = _a.sent();
                        if (!regionalResponse.ok) return [3 /*break*/, 6];
                        return [4 /*yield*/, regionalResponse.json()];
                    case 5:
                        data = _a.sent();
                        regionalChannelURL = data.channelUrlsById.directline;
                        return [3 /*break*/, 7];
                    case 6:
                        console.error("Regional settings error: ".concat(regionalResponse.status));
                        return [2 /*return*/];
                    case 7: return [4 /*yield*/, fetch(botURL)];
                    case 8:
                        response = _a.sent();
                        if (!response.ok) return [3 /*break*/, 10];
                        return [4 /*yield*/, response.json()];
                    case 9:
                        conversationInfo = _a.sent();
                        //console.log("Token for Direct Line:", conversationInfo.token);
                        //console.log("Direct Line domain:", `${regionalChannelURL}v3/directline`);
                        directline = createDirectLine({
                            token: conversationInfo.token,
                            domain: "".concat(regionalChannelURL, "v3/directline"),
                        });
                        // ✅ Save directline for Start Over button
                        setDirectlineInstance(directline);
                        return [3 /*break*/, 11];
                    case 10:
                        console.error("Bot token fetch failed: ".concat(response.status));
                        return [2 /*return*/];
                    case 11:
                        store = createStore({}, function (_a) {
                            var dispatch = _a.dispatch;
                            return function (next) {
                                return function (action) {
                                    var _a;
                                    if (props.greet &&
                                        action.type === "DIRECT_LINE/CONNECT_FULFILLED") {
                                        dispatch({
                                            meta: { method: "keyboard" },
                                            payload: {
                                                activity: {
                                                    channelData: { postBack: true },
                                                    name: "startConversation",
                                                    type: "event",
                                                },
                                            },
                                            type: "DIRECT_LINE/POST_ACTIVITY",
                                        });
                                    }
                                    if (action.type === "DIRECT_LINE/INCOMING_ACTIVITY") {
                                        var activity = action.payload.activity;
                                        if (((_a = activity.from) === null || _a === void 0 ? void 0 : _a.role) === "bot" &&
                                            getOAuthCardResourceUri(activity)) {
                                            directline
                                                .postActivity({
                                                type: "invoke",
                                                name: "signin/tokenExchange",
                                                value: {
                                                    id: activity.attachments[0].content.tokenExchangeResource
                                                        .id,
                                                    connectionName: activity.attachments[0].content.connectionName,
                                                    token: token,
                                                },
                                                from: {
                                                    id: props.userEmail,
                                                    name: props.userFriendlyName,
                                                    role: "user",
                                                },
                                            })
                                                .subscribe(function (id) {
                                                if (id === "retry")
                                                    return next(action);
                                            }, function (error) {
                                                console.error("OAuth invoke error:", error);
                                                return next(action);
                                            });
                                            return;
                                        }
                                    }
                                    return next(action);
                                };
                            };
                        });
                        canvasStyleOptions = {
                            hideUploadButton: false,
                        };
                        if (webChatRef.current && loadingSpinnerRef.current) {
                            webChatRef.current.style.minHeight = "50vh";
                            loadingSpinnerRef.current.style.display = "none";
                            renderWebChat({
                                directLine: directline,
                                store: store,
                                styleOptions: canvasStyleOptions,
                                userID: props.userEmail,
                            }, webChatRef.current);
                        }
                        return [2 /*return*/];
                }
            });
        }); };
        renderBot();
    }, [props]);
    // ✅ Start Over button click handler
    var handleStartOver = function () {
        if (directlineInstance) {
            directlineInstance
                .postActivity({
                type: "event",
                name: "StartOver", // Your bot must handle this
                from: { id: props.userEmail, name: props.userFriendlyName },
            })
                .subscribe(function (id) { return console.log("Start Over event sent:", id); }, function (error) {
                return console.error("Error sending Start Over event:", error);
            });
        }
    };
    // Till here
    return (React.createElement("div", { className: styles.chatContainer },
        React.createElement("div", { className: styles.chatHeader },
            React.createElement("div", { className: styles.chatAvatar }, props.botAvatarImage ? (React.createElement("img", { src: props.botAvatarImage, alt: "Bot Avatar" })) : ("🤖")),
            React.createElement("div", { className: styles.chatTitle }, props.botName || "Copilot Assistant"),
            React.createElement("button", { className: styles.startOverButton, onClick: handleStartOver }, "Start Over")),
        React.createElement("div", { ref: webChatRef, className: styles.webChat, role: "main" }),
        React.createElement("div", { ref: loadingSpinnerRef, className: styles.loadingSpinner },
            React.createElement(Spinner, { label: "Loading..." }))));
};
export default ServiceDeskChat;
//# sourceMappingURL=ServiceDeskChat.js.map