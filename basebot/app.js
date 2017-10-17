/* jshint esversion: 6 */
require('dotenv').config();
const restify = require('restify');
const fs = require('fs');
const builder = require('botbuilder');
const ticketsApi = require('./ticketsApi');
const azureSearch = require('./azureSearchApiClient');
const captions = require('./captions')

//configs
const listenPort = process.env.port || process.env.PORT || 3978;
const ticketSubmissionUrl = process.env.TICKET_SUBMISSION_URL || `http://localhost:${listenPort}`;

const azureSearchQuery = azureSearch({
    searchName: process.env.AZURE_SEARCH_ACCOUNT,
    indexName: process.env.AZURE_SEARCH_INDEX,
    searchKey: process.env.AZURE_SEARCH_KEY
});

// Setup Restify Server
const server = restify.createServer();
server.listen(listenPort, '::', () => {
    console.log('Server Up');
});

// Setup body parser and tickets api
server.use(restify.bodyParser());
server.post('/api/tickets', ticketsApi);

// Create chat connector for communicating with the Bot Framework Service
var connector = new builder.ChatConnector({
    appId: process.env.MICROSOFT_APP_ID,
    appPassword: process.env.MICROSOFT_APP_PASSWORD
});

// Listen for messages from users
server.post('/api/messages', connector.listen());

// exercise 3
var bot = new builder.UniversalBot(connector, (session) => {
    session.endDialog(`I'm sorry, I did not understand '${session.message.text}'.\nType 'help' to know more about me :)`);
});

// Connect to LUIS
// Be aware that the recognizer will run for every message received from the user.
var luisRecognizer = new builder.LuisRecognizer(process.env.LUIS_MODEL_URL).onEnabled((context, callback) => {
    var enabled = context.dialogStack().length === 0;
    callback(null, enabled);
});

bot.recognizer(luisRecognizer);

// reads from the cards template and replacing it with current data
const createCard = (ticketId, data) => {
    var cardTxt = fs.readFileSync('./cards/ticket.json', 'UTF-8');

    cardTxt = cardTxt.replace(/{ticketId}/g, ticketId)
                    .replace(/{severity}/g, data.severity)
                    .replace(/{category}/g, data.category)
                    .replace(/{description}/g, data.description);

    return JSON.parse(cardTxt);
};


bot.dialog('Help',
    (session, args, next) => { session.endDialog(captions.help); }
  ).triggerAction({
    matches: 'Help'
});


bot.dialog('SubmitTicket', [
    (session, args, next) => {
        var messages = {
          questionSeverity: 'Which is the severity of this problem?'
        }

        var category = builder.EntityRecognizer.findEntity(args.intent.entities, 'category');
        var severity = builder.EntityRecognizer.findEntity(args.intent.entities, 'severity');

        if (category && category.resolution.values.length > 0) {
            session.dialogData.category = category.resolution.values[0];
        }

        if (severity && severity.resolution.values.length > 0) {
            session.dialogData.severity = severity.resolution.values[0];
        }

        session.dialogData.description = session.message.text;

        if (!session.dialogData.severity) {
            var choices = ['high', 'normal', 'low'];
            builder.Prompts.choice(session,
              messages.questionSeverity,
              choices,
              { listStyle: builder.ListStyle.button }
            );
        } else {
            next();
        }
    },
    (session, result, next) => {
        var messages = {
          questionCategory: 'Which would be the category for this ticket (software, hardware, networking, security or other)?'
        }

        if (!session.dialogData.severity) {
            session.dialogData.severity = result.response.entity;
        }

        if (!session.dialogData.category) {
            builder.Prompts.text(session,
              messages.questionCategory
            );
        } else {
            next();
        }
    },
    (session, result, next) => {
        if (!session.dialogData.category) {
            session.dialogData.category = result.response;
        }

        var messages = {
          questionConfirmationCreateTicket: `Great! I'm going to create a "${session.dialogData.severity}" severity ticket in the "${session.dialogData.category}" category. ` +
                      `The description I will use is "${session.dialogData.description}". Can you please confirm that this information is correct?`
        }

        builder.Prompts.confirm(session,
          messages.questionConfirmationCreateTicket,
          { listStyle: builder.ListStyle.button }
        );
    },
    (session, result, next) => {
      // TODO
        var messages = {
          errorSavingTicket: 'Ooops! Something went wrong while I was saving your ticket. Please try again later.',
          ticketNotCreated: 'Ok. The ticket was not created. You can start again if you want.'
        }

        if (result.response) {
            var data = {
                category: session.dialogData.category,
                severity: session.dialogData.severity,
                description: session.dialogData.description,
            };

            const client = restify.createJsonClient({ url: ticketSubmissionUrl });

            client.post('/api/tickets', data, (err, request, response, ticketId) => {
                if (err || ticketId == -1) {
                    session.send(messages.errorSavingTicket);
                } else {
                    session.send(new builder.Message(session).addAttachment({
                        contentType: "application/vnd.microsoft.card.adaptive",
                        content: createCard(ticketId, data)
                    }));
                }

                session.endDialog();
            });
        } else {
            session.endDialog(messages.ticketNotCreated);
        }
    }
]).triggerAction({
    matches: 'SubmitTicket'
});

// execute a simple search using the text typed by the user
bot.dialog('SearchKB', [
    (session) => {
      var messages = {
        errorSearchingAzure: 'Ooops! Something went wrong while contacting Azure Search. Please try again later.'
      }

        session.sendTyping();
        azureSearchQuery(`search=${encodeURIComponent(session.message.text.substring('search about '.length))}`, (err, result) => {
            if (err) {
                session.send(messages.errorSearchingAzure);
                return;
            }
            session.replaceDialog('ShowKBResults', {
              result,
              originalText: session.message.text }
            );
        });
    }
])
.triggerAction({
    matches: /^search about (.*)/i
});


bot.dialog('ExploreKnowledgeBase', [
    (session, args, next) => {

        var messages = {
          errorConnectingAzure: 'Ooops! Something went wrong while contacting Azure Search. Please try again later.',
          promptWhichCategory: 'Let\'s see if I can find something in the knowledge base for you. Which category is your question about?'
        }

        var category = builder.EntityRecognizer.findEntity(args.intent.entities, 'category');

        if (!category) {
            // retrieve facets
            azureSearchQuery('facet=category', (error, result) => {
                if (error) {
                    session.endDialog(messages.errorConnectingAzure);
                } else {
                    var choices = result['@search.facets'].category.map(item=> `${item.value} (${item.count})`);
                    builder.Prompts.choice(session,
                      messages.promptWhichCategory,
                      choices,
                      { listStyle: builder.ListStyle.button }
                    );
                }
            });
        } else {
            if (!session.dialogData.category) {
                session.dialogData.category = category.entity;
            }

            next();
        }
    },
    (session, args) => {
      var messages = {
        errorContactingAzure: 'Ooops! Something went wrong while contacting Azure Search. Please try again later.'
      }

      var category = (session.diaglogData.category) ?
                        session.dialogData.category :
                        args.response.entity.replace(/\s\([^)]*\)/,'');

      // search by category
      azureSearchQuery('$filter=' + encodeURIComponent(`category eq '${category}'`), (error, result) => {
          if (error) {
              session.endDialog(messages.errorContactingAzure);
          } else {
              session.replaceDialog('ShowKBResults', { result, originalText: category });
          }
      });
    }
]).triggerAction({
    matches: 'ExploreKnowledgeBase'
});


bot.dialog('DetailsOf', [
    (session, args) => {
      var messages = {
        cantFindArticle: 'Sorry, I could not find that article.'
      }

        var title = session.message.text.substring('show me the article '.length);
        azureSearchQuery('$filter=' + encodeURIComponent(`title eq '${title}'`), (error, result) => {
            if (error || !result.value[0]) {
                session.endDialog(messages.cantFindArticle);
            } else {
                session.endDialog(result.value[0].text);
            }
        });
    }
]).triggerAction({
    matches: /^show me the article (.*)/i
});

bot.dialog('ShowKBResults', [
    (session, args) => {
        if (args.result.value.length > 0) {
            var msg = new builder.Message(session).attachmentLayout(builder.AttachmentLayout.carousel);
            args.result.value.forEach((faq, i) => {
                msg.addAttachment(
                    new builder.ThumbnailCard(session)
                        .title(faq.title)
                        .subtitle(`Category: ${faq.category} | Search Score: ${faq['@search.score']}`)
                        .text(faq.text.substring(0, Math.min(faq.text.length, 50) + '...'))
                        .images([builder.CardImage.create(session, 'https://raw.githubusercontent.com/GeekTrainer/help-desk-bot-lab/master/assets/botimages/head-smiling-medium.png')])
                        .buttons([{ title: 'More details', value: `show me the article ${faq.title}`, type: 'postBack' }])
                );
            });
            session.send(`These are some articles I\'ve found in the knowledge base for _'${args.originalText}'_, click **More details** to read the full article:`);
            session.endDialog(msg);
        } else {
            session.endDialog(`Sorry, I could not find any results in the knowledge base for _'${args.originalText}'_`);
        }
    }
]);
