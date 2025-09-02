import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
export interface IServiceDeskChatWebPartProps {
    botURL: string;
    botName?: string;
    buttonLabel?: string;
    botAvatarImage?: string;
    botAvatarInitials?: string;
    greet?: boolean;
    customScope: string;
    clientID: string;
    authority: string;
}
export default class ServiceDeskChatWebPart extends BaseClientSideWebPart<IServiceDeskChatWebPartProps> {
    private _environmentMessage;
    onInit(): Promise<void>;
    render(): void;
    private _getEnvironmentMessage;
    onDispose(): void;
    protected get dataVersion(): Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=ServiceDeskChatWebPart.d.ts.map