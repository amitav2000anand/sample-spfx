import * as React from "react";

import { Spinner } from "office-ui-fabric-react/lib/Spinner";
import { IServiceDeskChatProps } from "././IServiceDeskChatProps";
import MSALWrapper from "../../../utils/MSALWrapper";
import styles from "./ServiceDeskChat.module.scss";
import { sendChatHistoryBeacon, WebChatBuilder } from "../../../utils/helpers";

const ServiceDeskChat: React.FC<IServiceDeskChatProps> = (props) => {
  const webChatRef = React.useRef<HTMLDivElement>(null);
  const loadingSpinnerRef = React.useRef<HTMLDivElement>(null);
  const [directLineInstance, setDirectLineInstance] = React.useState<any>(null);

  const [msalToken, setMsalToken] = React.useState<string>("");
  const [kvToken, setKvToken] = React.useState<string>("");

  const botURL = props.botURL;
  const environmentEndPoint = botURL.slice(
    0,
    botURL.indexOf("/powervirtualagents"),
  );
  const apiVersion = botURL.slice(botURL.indexOf("api-version")).split("=")[1];
  const regionalChannelSettingsURL = `${environmentEndPoint}/powervirtualagents/regionalchannelsettings?api-version=${apiVersion}`;

  const onSessionClosed = (newDirectLine: any): void => {
    setDirectLineInstance(newDirectLine);
  };

  const renderBot = async (
    getKvTokenCallback: (token: string) => void,
  ): Promise<void> => {
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
    setMsalToken(responseToken?.accessToken ?? "");

    const keyvault_scope = "https://vault.azure.net/.default";
    let kvToken = await MSALWrapperInstance.handleLoggedInUser(
      [keyvault_scope],
      props.userEmail,
    );
    if (!kvToken) {
      kvToken = await MSALWrapperInstance.acquireAccessToken(
        [keyvault_scope],
        props.userEmail,
      );
    }
    if (getKvTokenCallback) {
      getKvTokenCallback(kvToken?.accessToken ?? "");
    }

    if (webChatRef.current && loadingSpinnerRef.current) {
      const builder = new WebChatBuilder()
        .setProps(props)
        .setMSALToken(responseToken?.accessToken ?? "")
        .setRegionalChannelSettingsURL(regionalChannelSettingsURL)
        .setBotURL(botURL)
        .setWebChatRef(webChatRef)
        .setLoadingSpinnerRef(loadingSpinnerRef);
      await builder.build(onSessionClosed);
    }
  };

  React.useEffect(() => {
    let token = "";
    renderBot((kvToken: string): void => {
      token = kvToken;
      setKvToken(token);
    });

    window.addEventListener("unload", () => sendChatHistoryBeacon(token));

    return () =>
      window.removeEventListener("unload", () => sendChatHistoryBeacon(token));
  }, [props]);

  const endDirectLine = (token: string) => {
    if (directLineInstance?.connectionStatus$) {
      directLineInstance.connectionStatus$.next(5);
    }
    if (directLineInstance?.activity$?.unsubscribe) {
      directLineInstance.activity$.unsubscribe();
    }

    sendChatHistoryBeacon(token);
  };

  const handleStartOver = async (): Promise<void> => {
    endDirectLine(kvToken);

    if (webChatRef.current && loadingSpinnerRef.current) {
      const builder = new WebChatBuilder()
        .setProps(props)
        .setMSALToken(msalToken)
        .setRegionalChannelSettingsURL(regionalChannelSettingsURL)
        .setBotURL(botURL)
        .setWebChatRef(webChatRef)
        .setLoadingSpinnerRef(loadingSpinnerRef);
      await builder.build(onSessionClosed);
    }
  };

  return (
    <div className={styles.chatContainer}>
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
