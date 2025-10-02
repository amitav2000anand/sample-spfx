import * as React from "react";

import { Spinner } from "office-ui-fabric-react/lib/Spinner";
import { useRef } from "react";
import styles from "./BubbleChatDialog.module.scss"; // ✅ SCSS file for styling
import { IconButton } from "office-ui-fabric-react/lib/Button";
import { IBubbleChatProps } from "./IBubbleChatProps";
import MSALWrapper from "../../../utils/MSALWrapper";
import { WebChatBuilder, sendChatHistoryBeacon } from "../../../utils/helpers";

export const BubbleChatDialog: React.FunctionComponent<IBubbleChatProps> = (
  props,
) => {
  const webChatRef = useRef<HTMLDivElement>(null);
  const loadingSpinnerRef = useRef<HTMLDivElement>(null);
  const [directLineInstance, setDirectLineInstance] = React.useState<any>(null);

  const [msalToken, setMsalToken] = React.useState<string>("");
  const [kvToken, setKvToken] = React.useState<string>("");
  const [isOpen, setIsOpen] = React.useState(false);

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

  const handleLayerDidMount = async (
    getKvTokenCallback?: (token: string) => void,
  ) => {
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
    setMsalToken(responseToken?.accessToken || "");

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
    window.addEventListener("unload", () => sendChatHistoryBeacon(kvToken));

    return () =>
      window.removeEventListener("unload", () =>
        sendChatHistoryBeacon(kvToken),
      );
  }, [kvToken]);

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
    <>
      {isOpen && (
        <div className={styles.chatDialog}>
          <div className={styles.chatHeader}>
            <div className={styles.headerLeft}>
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
            </div>

            <div className={styles.headerActions}>
              <button
                className={styles.startOverButton}
                onClick={handleStartOver}
              >
                Start Over
              </button>
              <button
                className={styles.closeButton}
                onClick={() => setIsOpen(false)}
              >
                ×
              </button>
            </div>
          </div>

          {/* Chat body */}
          <div className={styles.chatBody}>
            <div ref={webChatRef}></div>
            <div ref={loadingSpinnerRef} className={styles.loadingOverlay}>
              <Spinner label="Loading..." />
            </div>
          </div>
        </div>
      )}

      {/* Floating toggle button */}
      <IconButton
        className={styles.toggleButton}
        iconProps={{ iconName: "Chat" }}
        title="Chat Now"
        ariaLabel="Chat Now"
        onClick={() => {
          setIsOpen(!isOpen);
          if (!isOpen) handleLayerDidMount((kvToken) => setKvToken(kvToken));
        }}
      />
    </>
  );
};

export default class Chatbot extends React.Component<IBubbleChatProps> {
  constructor(props: IBubbleChatProps) {
    super(props);
  }
  public render(): JSX.Element {
    return <BubbleChatDialog {...this.props} />;
  }
}
