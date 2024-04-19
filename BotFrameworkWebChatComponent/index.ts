import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as WebChat from 'botframework-webchat';

const DEFAULT_LOCALE = 'en-US';

export class BotFrameworkWebChatComponent implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private _notifyOutputChanged: () => void;

    private _webChatContainer: HTMLDivElement | null = null;
    private _botTokenEndpoint: string = '';
    private _directLine: any = null;
    private _debug: boolean = true;

    constructor() { }

    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void {
        if(this._debug) console.log(`WebChatComponent:init()`);
        // Add control initialization code
        this._container = container;
        this._context = context;
        this._notifyOutputChanged = notifyOutputChanged;
        this._context.mode.trackContainerResize(true);
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        if(this._debug) console.log(`WebChatComponent:updateView()`);

        this.checkforBotChanges(context);
        this.checkForUIChanges(context);
    }

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void {}

    private checkforBotChanges(context: ComponentFramework.Context<IInputs>)
    {
        this._context = context;

        if(this._debug) console.log(`WebChatComponent:checkforBotChanges() check for changes to the bot token endpoint`);

        if(context.parameters.tokenEndpoint.raw !== this._botTokenEndpoint)
        {
            this._botTokenEndpoint = context.parameters.tokenEndpoint.raw || '';
            this._directLine = null;
            if(this._webChatContainer !== null)
            {
                this.endConversation();
                this._container.removeChild(this._webChatContainer);
                this._webChatContainer = null;
            }

            this.getBotToken(new URL(this._botTokenEndpoint))
                .then((token) => {
                    this._directLine = WebChat.createDirectLine({ token });
                    this.createWebChatContainer();
                    this.renderWebChat();
                    this.startConversation();
                });
        }
    }

    private checkForUIChanges(context: ComponentFramework.Context<IInputs>)
    {
        if(this._webChatContainer === null) return;

        var needsUpdate = false;

        if (context.mode.allocatedHeight !== this._webChatContainer.clientHeight || context.mode.allocatedWidth !== this._webChatContainer.clientWidth) 
        {
            needsUpdate = true;
        }

        if( context.parameters.fontSize.raw !== this._webChatContainer.style.fontSize)
        {
            needsUpdate = true;
        }

        if(needsUpdate)
        {
            if(this._debug) console.log(`WebChatComponent:checkForUIChanges() Updating web chat container`);
            this.updateWebChatContainer();
        }
    }

    private createWebChatContainer() 
    {
        if(this._webChatContainer !== null) return;

        if(this._debug) console.log(`WebChatComponent:createWebChatContainer() Creating web chat container`);

        this._webChatContainer = document.createElement('div');
        this._webChatContainer.id = "webchat";
        this._webChatContainer.role = 'main';
        this._webChatContainer.style.bottom = '0';
        this._webChatContainer.style.right = '0';
        this._webChatContainer.style.textAlign = 'left';
        this.updateWebChatContainer();
        this._container.appendChild(this._webChatContainer);
    }

    private updateWebChatContainer()
    {
        if(this._webChatContainer === null) return;

        this._webChatContainer.style.fontSize = this._context.parameters.fontSize.raw || '18px';
        this._webChatContainer.style.height = `${this._context.mode.allocatedHeight}px`;
        this._webChatContainer.style.width = `${this._context.mode.allocatedWidth}px`;
    }

    private async getBotToken(tokenEndpointURL: URL): Promise<string> {
        if (this._debug) console.log(`WebChatComponent:getBotToken() Getting DirectLine token for endpoint`);
        try {
            const response = await fetch(tokenEndpointURL.toString());
            if (!response.ok) {
                throw new Error('Failed to retrieve Direct Line token.');
            }
            const { token } = await response.json();
            if (this._debug) console.log(`WebChatComponent:getBotToken() Got bot token for endpoint ${tokenEndpointURL.toString()}`);
            return token;
        } catch (error) {
            console.error(`WebChatComponent:getBotToken() Error getting token direct line: ${JSON.stringify(error instanceof Error ? error.message : error)}`);
            return '';
        }
    }

    private renderWebChat() 
    {
        const styleSet = WebChat.createStyleSet(
        {
            accent: "#000000",
            hideUploadButton: true,
            backgroundColor: "#F8F8F8",
            bubbleBorderColor: "#f08040",
            sendBoxButtonColor: "#000000",
            transcriptTerminatorFontSize: "24px",
            timestampColor: "#f08040",
            rootHeight: '100%',
            rootWidth: '100%',
        });

        const avatarOptions = {
            botAvatarImage: "https://github.com/psycook/images/blob/main/Contoso%20Bank%20Chatbot.png?raw=true",
            botAvatarInitials: "HB",
            userAvatarImage: "https://github.com/psycook/images/blob/main/Contoso%20Bank%20Customer%20New.png?raw=true",
            userAvatarInitials: "MS",
        };

        WebChat.renderWebChat(
            {
                directLine: this._directLine,
                userID: 'Simon',
                username: 'simonco',
                locale: DEFAULT_LOCALE,
                styleSet,
                styleOptions: avatarOptions
            },
            document.getElementById('webchat')
        );
    }

    private startConversation() 
    {
        if(this._directLine === null) return;

        this._directLine.postActivity({
            from: { id: 'Simon', name: 'simonco' },
            name: 'startConversation',
            type: 'event'
        }).subscribe(
            (id:string) => console.log(`Posted activity, assigned ID ${id}`),
            (error:any) => console.log(`Error posting activity ${error}`)
        );
    }

    private endConversation() 
    {
        if(this._directLine !== null)
        {
            this._directLine.end();
        }
    }
}