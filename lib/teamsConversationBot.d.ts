import { TeamsActivityHandler, TurnContext } from 'botbuilder';
export declare class TeamsConversationBot extends TeamsActivityHandler {
    constructor();
    cardActivityAsync(context: TurnContext, isUpdate: any): Promise<void>;
    sendUpdateCard(context: TurnContext, cardActions: any): Promise<void>;
    sendWelcomeCard(context: TurnContext, cardActions: any): Promise<void>;
    getSingleMember(context: TurnContext): Promise<void>;
    mentionActivityAsync(context: TurnContext): Promise<void>;
    deleteCardActivityAsync(context: TurnContext): Promise<void>;
    messageAllMembersAsync(context: TurnContext): Promise<void>;
    getPagedMembers(context: TurnContext): Promise<any>;
}
