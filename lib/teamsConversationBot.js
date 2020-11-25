"use strict";
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
const botbuilder_1 = require("botbuilder");
const TextEncoder = require('util').TextEncoder;
class TeamsConversationBot extends botbuilder_1.TeamsActivityHandler {
    constructor() {
        super();
        this.onMessage((context, next) => __awaiter(this, void 0, void 0, function* () {
            botbuilder_1.TurnContext.removeRecipientMention(context.activity);
            const text = context.activity.text.trim().toLocaleLowerCase();
            if (text.includes('mention')) {
                yield this.mentionActivityAsync(context);
            }
            else if (text.includes('update')) {
                yield this.cardActivityAsync(context, true);
            }
            else if (text.includes('delete')) {
                yield this.deleteCardActivityAsync(context);
            }
            else if (text.includes('message')) {
                yield this.messageAllMembersAsync(context);
            }
            else if (text.includes('who')) {
                yield this.getSingleMember(context);
            }
            else {
                yield this.cardActivityAsync(context, false);
            }
            yield next();
        }));
        this.onTeamsMembersAddedEvent((membersAdded, teamInfo, context, next) => __awaiter(this, void 0, void 0, function* () {
            let newMembers = '';
            membersAdded.forEach((account) => {
                newMembers += account.id + ' ';
            });
            const name = !teamInfo ? 'not in team' : teamInfo.name;
            const card = botbuilder_1.CardFactory.heroCard('Account Added', `${newMembers} joined ${name}.`);
            const message = botbuilder_1.MessageFactory.attachment(card);
            yield context.sendActivity(message);
            yield next();
        }));
    }
    cardActivityAsync(context, isUpdate) {
        return __awaiter(this, void 0, void 0, function* () {
            const cardActions = [
                {
                    text: 'MessageAllMembers',
                    title: 'Message all members',
                    type: botbuilder_1.ActionTypes.MessageBack,
                    value: null
                },
                {
                    text: 'whoAmI',
                    title: 'Who am I?',
                    type: botbuilder_1.ActionTypes.MessageBack,
                    value: null
                },
                {
                    text: 'Delete',
                    title: 'Delete card',
                    type: botbuilder_1.ActionTypes.MessageBack,
                    value: null
                }
            ];
            if (isUpdate) {
                yield this.sendUpdateCard(context, cardActions);
            }
            else {
                yield this.sendWelcomeCard(context, cardActions);
            }
        });
    }
    sendUpdateCard(context, cardActions) {
        return __awaiter(this, void 0, void 0, function* () {
            const data = context.activity.value;
            data.count += 1;
            cardActions.push({
                text: 'UpdateCardAction',
                title: 'Update Card',
                type: botbuilder_1.ActionTypes.MessageBack,
                value: data
            });
            const card = botbuilder_1.CardFactory.heroCard('Updated card', `Update count: ${data.count}`, null, cardActions);
            // card.id = context.activity.replyToId;
            const message = botbuilder_1.MessageFactory.attachment(card);
            message.id = context.activity.replyToId;
            yield context.updateActivity(message);
        });
    }
    sendWelcomeCard(context, cardActions) {
        return __awaiter(this, void 0, void 0, function* () {
            const initialValue = {
                count: 0
            };
            cardActions.push({
                text: 'UpdateCardAction',
                title: 'Update Card',
                type: botbuilder_1.ActionTypes.MessageBack,
                value: initialValue
            });
            const card = botbuilder_1.CardFactory.heroCard('Welcome card', '', null, cardActions);
            yield context.sendActivity(botbuilder_1.MessageFactory.attachment(card));
        });
    }
    getSingleMember(context) {
        return __awaiter(this, void 0, void 0, function* () {
            let member;
            try {
                member = yield botbuilder_1.TeamsInfo.getMember(context, context.activity.from.id);
            }
            catch (e) {
                if (e.code === 'MemberNotFoundInConversation') {
                    context.sendActivity(botbuilder_1.MessageFactory.text('Member not found.'));
                    return;
                }
                else {
                    console.log(e);
                    throw e;
                }
            }
            const message = botbuilder_1.MessageFactory.text(`You are: ${member.name}`);
            yield context.sendActivity(message);
        });
    }
    mentionActivityAsync(context) {
        return __awaiter(this, void 0, void 0, function* () {
            const mention = {
                mentioned: context.activity.from,
                text: `<at>${new TextEncoder().encode(context.activity.from.name)}</at>`,
                type: 'mention'
            };
            const replyActivity = botbuilder_1.MessageFactory.text(`Hi ${mention.text}`);
            replyActivity.entities = [mention];
            yield context.sendActivity(replyActivity);
        });
    }
    deleteCardActivityAsync(context) {
        return __awaiter(this, void 0, void 0, function* () {
            yield context.deleteActivity(context.activity.replyToId);
        });
    }
    // If you encounter permission-related errors when sending this message, see
    // https://aka.ms/BotTrustServiceUrl
    messageAllMembersAsync(context) {
        return __awaiter(this, void 0, void 0, function* () {
            const members = yield this.getPagedMembers(context);
            members.forEach((teamMember) => __awaiter(this, void 0, void 0, function* () {
                console.log('a ', teamMember);
                const message = botbuilder_1.MessageFactory.text(`Hello ${teamMember.givenName} ${teamMember.surname}. I'm a Teams conversation bot.`);
                const ref = botbuilder_1.TurnContext.getConversationReference(context.activity);
                ref.user = teamMember;
                let botAdapter;
                botAdapter = context.adapter;
                yield botAdapter.createConversation(ref, (t1) => __awaiter(this, void 0, void 0, function* () {
                    const ref2 = botbuilder_1.TurnContext.getConversationReference(t1.activity);
                    yield t1.adapter.continueConversation(ref2, (t2) => __awaiter(this, void 0, void 0, function* () {
                        yield t2.sendActivity(message);
                    }));
                }));
            }));
            yield context.sendActivity(botbuilder_1.MessageFactory.text('All messages have been sent.'));
        });
    }
    getPagedMembers(context) {
        return __awaiter(this, void 0, void 0, function* () {
            let continuationToken;
            const members = [];
            do {
                const pagedMembers = yield botbuilder_1.TeamsInfo.getPagedMembers(context, 100, continuationToken);
                continuationToken = pagedMembers.continuationToken;
                members.push(...pagedMembers.members);
            } while (continuationToken !== undefined);
            return members;
        });
    }
}
exports.TeamsConversationBot = TeamsConversationBot;
//# sourceMappingURL=teamsConversationBot.js.map