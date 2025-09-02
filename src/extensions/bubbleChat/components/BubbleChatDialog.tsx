import * as React from "react";
import * as ReactWebChat from "botframework-webchat";
import { Spinner } from "office-ui-fabric-react/lib/Spinner";
import { Dispatch } from "redux";
import { useRef } from "react";
import styles from "./BubbleChatDialog.module.scss"; // ✅ SCSS file for styling
import { IconButton } from "office-ui-fabric-react/lib/Button";
import { IBubbleChatProps } from "./IBubbleChatProps";
import MSALWrapper from "../../../utils/MSALWrapper";

export const BubbleChatDialog: React.FunctionComponent<IBubbleChatProps> = (
  props,
) => {

  //const [hideDialog, { toggle: toggleHideDialog }] = useBoolean(true);
  const [isOpen, setIsOpen] = React.useState(false);

  // Your bot's token endpoint
  const botURL = props.botURL;

  // constructing URL using regional settings
  const environmentEndPoint = botURL.slice(
    0,
    botURL.indexOf("/powervirtualagents"),
  );
  const apiVersion = botURL.slice(botURL.indexOf("api-version")).split("=")[1];
  const regionalChannelSettingsURL = `${environmentEndPoint}/powervirtualagents/regionalchannelsettings?api-version=${apiVersion}`;

  // Using refs instead of IDs to get the webchat and loading spinner elements
  const webChatRef = useRef<HTMLDivElement>(null);
  const loadingSpinnerRef = useRef<HTMLDivElement>(null);

  // A utility function that extracts the OAuthCard resource URI from the incoming activity or return undefined
  function getOAuthCardResourceUri(activity: any): string | undefined {
    const attachment = activity?.attachments?.[0];
    if (
      attachment?.contentType === "application/vnd.microsoft.card.oauth" &&
      attachment.content.tokenExchangeResource
    ) {
      return attachment.content.tokenExchangeResource.uri;
    }
  }

  const handleLayerDidMount = async () => {
    const MSALWrapperInstance = new MSALWrapper(
      props.clientID,
      props.authority,
    );

    // Trying to get token if user is already signed-in
    let responseToken = await MSALWrapperInstance.handleLoggedInUser(
      [props.customScope],
      props.userEmail,
    );

    if (!responseToken) {
      // Trying to get token if user is not signed-in
      responseToken = await MSALWrapperInstance.acquireAccessToken(
        [props.customScope],
        props.userEmail,
      );
    }

    const token = responseToken?.accessToken || null;

    // Get the regional channel URL
    let regionalChannelURL;

    const regionalResponse = await fetch(regionalChannelSettingsURL);
    if (regionalResponse.ok) {
      const data = await regionalResponse.json();
      regionalChannelURL = data.channelUrlsById.directline;
    } else {
      console.error(`HTTP error! Status: ${regionalResponse.status}`);
    }

    // Create DirectLine object
    let directline: any;

    const response = await fetch(botURL);

    if (response.ok) {
      const conversationInfo = await response.json();
      directline = ReactWebChat.createDirectLine({
        token: conversationInfo.token,
        domain: regionalChannelURL + "v3/directline",
      });
    } else {
      console.error(`HTTP error! Status: ${response.status}`);
    }

    const store = ReactWebChat.createStore(
      {},
      ({ dispatch }: { dispatch: Dispatch }) =>
        (next: any) =>
          (action: any) => {
            // Checking whether we should greet the user
            if (props.greet) {
              if (action.type === "DIRECT_LINE/CONNECT_FULFILLED") {
                console.log("Action:" + action.type);
                dispatch({
                  meta: {
                    method: "keyboard",
                  },
                  payload: {
                    activity: {
                      channelData: {
                        postBack: true,
                      },
                      //Web Chat will show the 'Greeting' System Topic message which has a trigger-phrase 'hello'
                      name: "startConversation",
                      type: "event",
                    },
                  },
                  type: "DIRECT_LINE/POST_ACTIVITY",
                });
                return next(action);
              }
            }

            // Checking whether the bot is asking for authentication
            if (action.type === "DIRECT_LINE/INCOMING_ACTIVITY") {
              const activity = action.payload.activity;
              if (
                activity.from &&
                activity.from.role === "bot" &&
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
                      if (id === "retry") {
                        // bot was not able to handle the invoke, so display the oauthCard (manual authentication)
                        console.log(
                          "bot was not able to handle the invoke, so display the oauthCard",
                        );
                        return next(action);
                      }
                    },
                    (error: any) => {
                      // an error occurred to display the oauthCard (manual authentication)
                      console.log("An error occurred so display the oauthCard");
                      return next(action);
                    },
                  );
                // token exchange was successful, do not show OAuthCard
                return;
              }
            } else {
              return next(action);
            }

            return next(action);
          },
    );

    // hide the upload button - other style options can be added here
    const canvasStyleOptions = {
      hideUploadButton: false,
    };

    // Render webchat
    if (token && directline) {
      if (webChatRef.current && loadingSpinnerRef.current) {
        webChatRef.current.style.minHeight = "50vh";
        loadingSpinnerRef.current.style.display = "none";
        ReactWebChat.renderWebChat(
          {
            directLine: directline,
            store,
            styleOptions: canvasStyleOptions,
            userID: props.userEmail,
          },
          webChatRef.current,
        );
      } else {
        console.error("Webchat or loading spinner not found");
      }
    }
  };

return (
  <>
    {isOpen && (
      <div className={styles.chatDialog}>
        {/* Header */}
        <div className={styles.chatHeader}>
          <span>{props.botName}</span>
          <button
            className={styles.closeButton}
            onClick={() => setIsOpen(false)}
          >
            ×
          </button>
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
        if (!isOpen) handleLayerDidMount();
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
    return (
      <BubbleChatDialog {...this.props} />
    );
  }
}
