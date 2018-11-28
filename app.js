/*-----------------------------------------------------------------------------
A simple Language Understanding (LUIS) bot for the Microsoft Bot Framework. 
-----------------------------------------------------------------------------*/

var request = require('request')
var restify = require('restify');
var builder = require('botbuilder');
var botbuilder_azure = require("botbuilder-azure");

// Setup Restify Server
var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function () {
   console.log('%s listening to %s', server.name, server.url); 
});
  
// Create chat connector for communicating with the Bot Framework Service
var connector = new builder.ChatConnector({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword,
    openIdMetadata: process.env.BotOpenIdMetadata 
});

// Listen for messages from users 
server.post('/api/messages', connector.listen());

/*----------------------------------------------------------------------------------------
* Bot Storage: This is a great spot to register the private state storage for your bot. 
* We provide adapters for Azure Table, CosmosDb, SQL Azure, or you can implement your own!
* For samples and documentation, see: https://github.com/Microsoft/BotBuilder-Azure
* ---------------------------------------------------------------------------------------- */

var tableName = 'botdata';
var azureTableClient = new botbuilder_azure.AzureTableClient(tableName, process.env['AzureWebJobsStorage']);
var tableStorage = new botbuilder_azure.AzureBotStorage({ gzipData: false }, azureTableClient);

// Create your bot with a function to receive messages from the user
// This default message handler is invoked if the user's utterance doesn't
// match any intents handled by other dialogs.
var bot = new builder.UniversalBot(connector,function (session) {
    //initialize conversational data for every new conversation
   session.send('I\'m sorry, I did not understand that');

     // If the object for storing notes in session.userData doesn't exist yet, initialize it
   if (!session.userData.shoppingCarts) {
       session.userData.shoppingCarts = {};
       console.log("initializing userData.notes in default message handler");
   }
  
 }

);

bot.set('storage', tableStorage);

// Make sure you add code to validate these fields
var luisAppId = process.env.LuisAppId;
var luisAPIKey = process.env.LuisAPIKey;
var luisAPIHostName = process.env.LuisAPIHostName || 'westus.api.cognitive.microsoft.com';

const LuisModelUrl = 'https://' + luisAPIHostName + '/luis/v2.0/apps/' + luisAppId + '?subscription-key=' + luisAPIKey;

// Create a recognizer that gets intents from LUIS, and add it to the bot
var recognizer = new builder.LuisRecognizer(LuisModelUrl);
bot.recognizer(recognizer);

// Add a dialog for each intent that the LUIS app recognizes.
// See https://docs.microsoft.com/en-us/bot-framework/nodejs/bot-builder-nodejs-recognize-intent-luis 
bot.dialog('GreetingDialog',[
  //(session) => {builder.Prompts.text(session, "Hi...How can I help ?")},
  (session) => {
    //session.dialogData.firstInput = results.response
    // make the api call here with the inputs received from the user
    // below example is for a post call
      builder.Prompts.text(session, "Hi...How can I help ?")
      request.post('http://104.211.102.154:3333/del_cart', {
        'auth': {
            'user': 'abc',
            'pass': 'xyz',
            'sendImmediately': false
          }, 
          'json': {
            //item: session.dialogData.firstInput
            item: ""
          }
        }, (error, response, body) => {
                var data = body;
                if (data['status'] == "successful") {
        		session.endDialog();
    			}
                else {
                        session.send('We are sorry, we are having system issue, please start over', data.item);
        		        session.endDialog();
    			     };
              })        
  }
]).triggerAction({
    matches: 'Greeting'
});




// Dialog or Find Items.
bot.dialog('Shopping.FindItemDialog', [
  (session) => {builder.Prompts.text(session, "Sure ! What do you like to buy ?")},
  (session, results) => {
    session.dialogData.firstInput = results.response
    // make the api call here with the inputs received from the user
    // below example is for a post call
      request.post('http://104.211.102.154:3333/search', {
        'auth': {
            'user': 'abc',
            'pass': 'xyz',
            'sendImmediately': false
          }, 
          'json': {
            item: session.dialogData.firstInput
          }
        }, (error, response, body) => {
                var data = body;
                if (data['status'] == "found") {
        		session.send('\%s\ is available with us, price is Rs. \%s\ ', data['item'], data['price']);
                //session.send('Price is : Rs. \%s\ ', data['price']);
        		session.endDialog();
    			}
                else {
                        session.send('We are sorry, item is not available with us today', data.item);
        		        session.endDialog();
    			     };
              })        
  }
]
).triggerAction({
    matches: 'Shopping.FindItem'
});





//  dialog for Shopping Cart Adding
bot.dialog('AddToShoppingCartDialog', [
    function (session, args, next) {
        // Resolve and store any Note.shoppingItem entity passed from LUIS.
        var intent = args.intent;
        var shoppingItem = builder.EntityRecognizer.findEntity(intent.entities, 'Shopping.Item');
        var shoppingQuantity = builder.EntityRecognizer.findEntity(intent.entities, 'Shopping.Quantity');
        var shoppingCart = session.dialogData.shoppingCart = {
          shoppingItem: shoppingItem ? shoppingItem.entity : null,
        };
        
        // Prompt for shoppingItem
        if (!shoppingCart.shoppingItem) {
            builder.Prompts.text(session, 'Which item you want to shop?');
        } else {
            next();
        }
    },
    function (session, results, next) {
        var shoppingCart = session.dialogData.shoppingCart;
        if (results.response) {
            shoppingCart.shoppingItem = results.response;
        }

        // Prompt for the text of the note
        if (!shoppingCart.shoppingQuantity) {
            builder.Prompts.text(session, 'What quantity would you like  to buy?');
        } else {
            next();
        }
    },
    function (session, results) {
        var shoppingCart = session.dialogData.shoppingCart;
        if (results.response) {
            shoppingCart.shoppingQuantity = results.response;
        }
        
        // If the object for storing notes in session.userData doesn't exist yet, initialize it
        if (!session.userData.shoppingCarts) {
            session.userData.shoppingCarts = {};
            console.log("initializing session.userData.shoppingCart in AddToCart dialog");
        }
        // Save notes in the notes object
        session.userData.shoppingCarts[shoppingCart.shoppingItem] = shoppingCart;

        // Send confirmation to user

        request.post('http://104.211.102.154:3333/cart', {
        'auth': {
            'user': 'abc',
            'pass': 'xyz',
            'sendImmediately': false
          }, 
          'json': {
            shopping_item: shoppingCart.shoppingItem,
            shopping_quantity: shoppingCart.shoppingQuantity 
          }
        }, (error, response, body) => {
                var data = body;
                if (data['status'] == "successful") {
                        session.endDialog('Item "%s" with quantity of "%s added to your shopping cart"',
                        shoppingCart.shoppingItem, shoppingCart.shoppingQuantity);
    			}
                else {
                        session.endDialog('Technical problem, items cannot be added, please try again later');
    			     };
              })
    }
]).triggerAction({ 
    matches: 'Shopping.AddToCart',
    confirmPrompt: "This will cancel the creation of the note you started. Are you sure?" 
}).cancelAction('cancelAddToCart', "Add To Cart canceled.", {
    matches: /^(cancel|nevermind)/i,
    confirmPrompt: "Are you sure?"
});




// Dialog or Buy Items.
bot.dialog('Shopping.BuyItemDialog',[
  (session) => {builder.Prompts.text(session, "Do you like to process your cart ?, please confirm")},
  (session, results) => {
    session.dialogData.firstInput = results.response
    // make the api call here with the inputs received from the user
    // below example is for a post call
      request.post('http://104.211.102.154:3333/show_cart', {
        'auth': {
            'user': 'abc',
            'pass': 'xyz',
            'sendImmediately': false
          }, 
          'json': {
            item: session.dialogData.firstInput
          }
        }, (error, response, body) => {
                var data = body;
                if (data['status'] == "found_cart") {
                        session.send('Your cart total is : Rs. \%s\ ', data['cart_total']);
        		session.endDialog();
    			}
                else {
                        session.send('You do not have any item in the cart');
        		        session.endDialog();
    			     };
              })        
  }
]
).triggerAction({
    matches: 'Shopping.BuyItem'
});

