import {
  createDirectLine,
  createStore,
  renderWebChat,
} from "botframework-webchat";
import * as React from "react";
import { IServiceDeskChatProps } from "../webparts/serviceDeskChat/components/IServiceDeskChatProps";
import { Dispatch, Store } from "redux";
import * as MarkdownIt from "markdown-it";

const getOAuthCardResourceUri = (activity: any): string | undefined => {
  const attachment = activity?.attachments?.[0];
  if (
    attachment?.contentType === "application/vnd.microsoft.card.oauth" &&
    attachment.content.tokenExchangeResource
  ) {
    return attachment.content.tokenExchangeResource.uri;
  }
};

const isAdaptiveCard = (activity: any) => {
  return (
    activity?.attachments?.[0]?.contentType ===
    "application/vnd.microsoft.card.adaptive"
  );
};

export const isSessionClosed = (activity: any): boolean => {
  if (!activity) return false;

  // This is fragile. If the bot response changes, this will break.
  return (
    activity.text &&
    activity.text.toLowerCase() ===
      "your session has been closed due to inactivity."
  );
};

const createOrAppendChatHistory = (activity: any): void => {
  if (!activity) return;

  if (activity.type !== "message") return;

  if (getOAuthCardResourceUri(activity)) return;

  if (isAdaptiveCard(activity)) return;

  if (isSessionClosed(activity)) return;

  const rawHistory = localStorage.getItem("sdChatHistory");
  const initialHistory = {
    conversationId: activity.conversation.id,
    channelId: activity.channelId,
    messages: [],
  };
  const sdChatHistory = rawHistory
    ? JSON.parse(rawHistory).conversationId !== activity.conversation.id
      ? initialHistory
      : JSON.parse(rawHistory)
    : initialHistory;

  sdChatHistory.messages.push({
    id: activity.from.id,
    name: activity.from.name,
    role: activity.from.role,
    message: activity.text,
  });

  localStorage.setItem("sdChatHistory", JSON.stringify(sdChatHistory));
};

const getEnv = () => {
  // This is fragile. make sure we stick to the same local, dev or stage site
  const pathname = window.location.pathname.toLowerCase();

  if (pathname.includes("workbench.aspx")) return "local";
  if (pathname.includes("service-desk-chatbot-dev.aspx")) return "dev";
  if (pathname.includes("service-desk-chatbot.aspx")) return "stage";
  return "prod";
};

const getUrl = (): string => {
  if (getEnv() === "local") return "http://localhost:7071/api/chat-history";

  return `https://essaiapimanagementservice-${getEnv()}.azure-api.net/employee-self-service/v1/chat-history`;
};

// const getKeyvaultName = (): string => {
//   if (getEnv() === "local") return "";
//   return `ess-ai-kv-${getEnv()}`;
// };

const getApimSubscriptionKey = async (
  token?: string,
): Promise<string | undefined> => {
  return new Promise((resolve) => {
    if (getEnv() === "local") {
      resolve("");
    } else if (getEnv() === "dev") {
      resolve("");
    } else if (getEnv() === "stage") {
      resolve("");
    } else {
      resolve("");
    }
  });

  // const url = `https://${getKeyvaultName()}.vault.azure.net`;

  // const credential = {
  //   getToken: async () => ({
  //     token,
  //     expiresOnTimestamp: Date.now() + 3600 * 1000,
  //   }),
  // };
  // const client = new SecretClient(url, credential);
  // const secretName = "AIApiManagementSubscriptionKey";
  // try {
  //   const result = await client.getSecret(secretName);
  //   return result.value;
  // } catch (error) {
  //   console.error("KeyVault error:", error);
  //   return undefined;
  // }
};

export const sendChatHistoryBeacon = (token?: string): void => {
  const rawHistory = localStorage.getItem("sdChatHistory");
  if (!rawHistory) return;

  getApimSubscriptionKey(token).then((key: string) => {
    fetch(getUrl(), {
      body: rawHistory,
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: key,
      },
    });
  });
};

const createDirectLineInstance = async (
  botURL: string,
  directLineDomain: string,
) => {
  try {
    const response = await fetch(botURL);
    const conversationInfo = await response.json();
    const directLine = createDirectLine({
      token: conversationInfo.token,
      domain: `${directLineDomain}v3/directline`,
    });
    return directLine;
  } catch (error) {
    console.error("DirectLine error:", error);
    return undefined;
  }
};

export const createBotStore = (
  props: IServiceDeskChatProps,
  token: string,
  directline: any,
  onSessionClosed: () => void,
): Store<any, any> => {
  return createStore(
    {},
    ({ dispatch }: { dispatch: Dispatch }) =>
      (next: any) =>
      (action: any) => {
        if (props.greet && action.type === "DIRECT_LINE/CONNECT_FULFILLED") {
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
          const activity = action.payload.activity;
          if (
            activity.from?.role === "bot" &&
            getOAuthCardResourceUri(activity)
          ) {
            directline
              .postActivity({
                type: "invoke",
                name: "signin/tokenExchange",
                value: {
                  id: activity.attachments[0].content.tokenExchangeResource.id,
                  connectionName:
                    activity.attachments[0].content.connectionName,
                  token,
                },
                from: {
                  id: props.userEmail,
                  name: props.userFriendlyName,
                  role: "user",
                },
              })
              .subscribe(
                (id: any) => {
                  if (id === "retry") return next(action);
                },
                (error: any) => {
                  console.error("OAuth invoke error:", error);
                  return next(action);
                },
              );
            return;
          }
          if (isSessionClosed(activity)) {
            if (onSessionClosed) {
              onSessionClosed();
            }
            return;
          }
          createOrAppendChatHistory(activity);
        }
        return next(action);
      },
  );
};

export const renderMarkdown = (text: string): string => {
  const md = new (MarkdownIt as any).default({
    html: false,
    linkify: true,
  });

  const defaultRender =
    md.renderer.rules.link_open ||
    ((tokens: any[], idx: number, options: any, env: any, self: any): string =>
      self.renderToken(tokens, idx, options));

  md.renderer.rules.link_open = (
    tokens: any[],
    idx: number,
    options: any,
    env: any,
    self: any,
  ): string => {
    const token = tokens[idx];
    const hrefIndex = token.attrIndex("href");
    const href = hrefIndex >= 0 ? token.attrs![hrefIndex][1] : "";

    // Add target="_blank"
    const targetIndex = token.attrIndex("target");
    if (targetIndex < 0) {
      token.attrPush(["target", "_blank"]);
    } else {
      token.attrs![targetIndex][1] = "_blank";
    }

    // Add rel="noopener noreferrer"
    const relIndex = token.attrIndex("rel");
    if (relIndex < 0) {
      token.attrPush(["rel", "noopener noreferrer"]);
    } else {
      token.attrs![relIndex][1] = "noopener noreferrer";
    }

    // Optional: Add title
    const titleIndex = token.attrIndex("title");
    if (titleIndex < 0) {
      token.attrPush(["title", "Opens in a new tab"]);
    }

    // see the reference: https://learn.microsoft.com/en-us/sharepoint/dev/spfx/hyperlinking
    let isSharePoint = false;
    try {
      const urlObj = new URL(href, "http://dummy.base"); // 'dummy.base' allows relative URLs
      const host = urlObj.host.toLowerCase();
      if (host === "sharepoint.com" || host.endsWith(".sharepoint.com")) {
        isSharePoint = true;
      }
    } catch (e) {
      console.log(e);
      // If URL parsing fails, treat as not SharePoint
    }
    if (isSharePoint) {
      const interceptionIndex = token.attrIndex("data-interception");
      if (interceptionIndex < 0) {
        token.attrPush(["data-interception", "off"]);
      } else {
        token.attrs![interceptionIndex][1] = "off";
      }
    }

    const linkOpenTag = defaultRender(tokens, idx, options, env, self);
    // This is a 1x1 transparent pixel gif as a placeholder for an icon. Default icon used by Bot Framework Web Chat.
    const iconHtml = `<img src="data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7" alt="Opens in a new window; external." class="webchat__markdown__external-link-icon">&nbsp;`;
    return linkOpenTag + iconHtml;
  };

  return md.render(text);
};

export class WebChatBuilder {
  private regionalChannelSettingsURL: string = "";
  private botURL: string = "";
  private props: any | undefined = undefined;
  private msalToken: string = "";
  private styleOptions: any = {
    hideUploadButton: false,
  };
  private webChatRef: React.RefObject<HTMLDivElement> = React.createRef();
  private loadingSpinnerRef: React.RefObject<HTMLDivElement> =
    React.createRef();

  public setProps(props: { [key: string]: any }): WebChatBuilder {
    this.props = props;
    return this;
  }

  public setRegionalChannelSettingsURL(url: string) {
    this.regionalChannelSettingsURL = url;
    return this;
  }

  public setBotURL(url: string): WebChatBuilder {
    this.botURL = url;
    return this;
  }

  public setMSALToken(token: string): WebChatBuilder {
    this.msalToken = token;
    return this;
  }

  public setStyleOptions(options: { [key: string]: any }): WebChatBuilder {
    this.styleOptions = options;
    return this;
  }

  public setWebChatRef(ref: React.RefObject<HTMLDivElement>): WebChatBuilder {
    this.webChatRef = ref;
    return this;
  }

  public setLoadingSpinnerRef(
    ref: React.RefObject<HTMLDivElement>,
  ): WebChatBuilder {
    this.loadingSpinnerRef = ref;
    return this;
  }

  public async build(onSessionClosed?: (newDirectLine: any) => void) {
    try {
      if (!this.webChatRef.current || !this.loadingSpinnerRef.current)
        throw new Error("WebChatBuilder: refs are not set");

      if (!this.props) throw new Error("WebChatBuilder: props are not set");

      const regionalResponse = await fetch(this.regionalChannelSettingsURL);
      const data = await regionalResponse.json();
      const regionalChannelURL = data.channelUrlsById.directline;

      const directLine = await createDirectLineInstance(
        this.botURL,
        regionalChannelURL,
      );

      const store = createBotStore(
        this.props,
        this.msalToken,
        directLine,
        async () => {
          const builder = new WebChatBuilder()
            .setProps(this.props)
            .setMSALToken(this.msalToken)
            .setRegionalChannelSettingsURL(this.regionalChannelSettingsURL)
            .setBotURL(this.botURL)
            .setWebChatRef(this.webChatRef)
            .setLoadingSpinnerRef(this.loadingSpinnerRef);
          await builder.build(onSessionClosed);
        },
      );

      this.webChatRef.current.style.minHeight = "50vh";
      this.loadingSpinnerRef.current.style.display = "none";

      renderWebChat(
        {
          directLine,
          store,
          styleOptions: this.styleOptions,
          userID: this.props.userEmail,
          username: this.props.userFriendlyName,
          renderMarkdown: renderMarkdown,
        },
        this.webChatRef.current,
      );

      if (onSessionClosed) {
        onSessionClosed(directLine);
      }
    } catch (error) {
      console.error("WebChatBuilder error:", error);
      return undefined;
    }
  }
}
