import * as React from "react";
import { useState } from "react";
import {
  createDirectLine,
  renderWebChat,
  createStore,
} from "botframework-webchat";

import { Spinner } from "office-ui-fabric-react/lib/Spinner";
import { Dispatch } from "redux";
import { useRef, useEffect } from "react";
import { IServiceDeskChatProps } from "././IServiceDeskChatProps";
import MSALWrapper from "../../../utils/MSALWrapper";
import styles from "./ServiceDeskChat.module.scss";

const ServiceDeskChat: React.FC<IServiceDeskChatProps> = (props) => {
  const webChatRef = useRef<HTMLDivElement>(null);
  const loadingSpinnerRef = useRef<HTMLDivElement>(null);

  // ✅ Store directline so we can send Start Over event
  const [directlineInstance, setDirectlineInstance] = useState<any>(null);
  const botURL = props.botURL;

  const environmentEndPoint = botURL.slice(
    0,
    botURL.indexOf("/powervirtualagents"),
  );
  const apiVersion = botURL.slice(botURL.indexOf("api-version")).split("=")[1];
  const regionalChannelSettingsURL = `${environmentEndPoint}/powervirtualagents/regionalchannelsettings?api-version=${apiVersion}`;

  const getOAuthCardResourceUri = (activity: any): string | undefined => {
    const attachment = activity?.attachments?.[0];
    if (
      attachment?.contentType === "application/vnd.microsoft.card.oauth" &&
      attachment.content.tokenExchangeResource
    ) {
      return attachment.content.tokenExchangeResource.uri;
    }
  };

  useEffect(() => {
    const renderBot = async (): Promise<void> => {
      const MSALWrapperInstance = new MSALWrapper(
        props.clientID,
        props.authority,
      );

      let responseToken = await MSALWrapperInstance.handleLoggedInUser(
        [props.customScope],
        props.userEmail,
      );
      if (!responseToken) {
        responseToken = await MSALWrapperInstance.acquireAccessToken(
          [props.customScope],
          props.userEmail,
        );
      }

      const token = responseToken?.accessToken || null;

      let regionalChannelURL;
      const regionalResponse = await fetch(regionalChannelSettingsURL);
      if (regionalResponse.ok) {
        const data = await regionalResponse.json();
        regionalChannelURL = data.channelUrlsById.directline;
      } else {
        console.error(`Regional settings error: ${regionalResponse.status}`);
        return;
      }

      let directline: any;
      const response = await fetch(botURL);
      if (response.ok) {
        const conversationInfo = await response.json();
        //console.log("Token for Direct Line:", conversationInfo.token);
        //console.log("Direct Line domain:", `${regionalChannelURL}v3/directline`);

        directline = createDirectLine({
          token: conversationInfo.token,
          domain: `${regionalChannelURL}v3/directline`,
        });
        // ✅ Save directline for Start Over button
        setDirectlineInstance(directline);
      } else {
        console.error(`Bot token fetch failed: ${response.status}`);
        return;
      }

      const store = createStore(
        {},
        ({ dispatch }: { dispatch: Dispatch }) =>
          (next: any) =>
          (action: any) => {
            if (
              props.greet &&
              action.type === "DIRECT_LINE/CONNECT_FULFILLED"
            ) {
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
                      id: activity.attachments[0].content.tokenExchangeResource
                        .id,
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
            }

            return next(action);
          },
      );

      const canvasStyleOptions = {
        hideUploadButton: false,
      };

      if (webChatRef.current && loadingSpinnerRef.current) {
        webChatRef.current.style.minHeight = "50vh";
        loadingSpinnerRef.current.style.display = "none";

        renderWebChat(
          {
            directLine: directline,
            store,
            styleOptions: canvasStyleOptions,
            userID: props.userEmail,
          },
          webChatRef.current,
        );
      }
    };

    renderBot();
  }, [props]);

  // ✅ Start Over button click handler
  const handleStartOver = (): void => {
    if (directlineInstance) {
      directlineInstance
        .postActivity({
          type: "event",
          name: "StartOver", // Your bot must handle this
          from: { id: props.userEmail, name: props.userFriendlyName },
        })
        .subscribe(
          (id: any) => console.log("Start Over event sent:", id),
          (error: any) =>
            console.error("Error sending Start Over event:", error),
        );
    }
  };
  // Till here
  return (
    <div className={styles.chatContainer}>
      {/* ✅ Single green header */}
      <div className={styles.chatHeader}>
        <div className={styles.chatAvatar}>
          {props.botAvatarImage ? (
            <img src={props.botAvatarImage} alt="Bot Avatar" />
          ) : (
            "🤖"
          )}
        </div>
        <div className={styles.chatTitle}>
          {props.botName || "Copilot Assistant"}
        </div>
        <button className={styles.startOverButton} onClick={handleStartOver}>
          Start Over
        </button>
      </div>

      <div ref={webChatRef} className={styles.webChat} role="main" />
      <div ref={loadingSpinnerRef} className={styles.loadingSpinner}>
        <Spinner label="Loading..." />
      </div>
    </div>
  );
};
export default ServiceDeskChat;
