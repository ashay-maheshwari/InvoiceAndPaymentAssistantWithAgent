/*-----------------------------------------------------------------------------
A simple Language Understanding (LUIS) bot for the Microsoft Bot Framework. 
-----------------------------------------------------------------------------*/

var restify = require('restify');
var builder = require('botbuilder');
const request = require("request");
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
var bot = new builder.UniversalBot(connector);
bot.set('storage', tableStorage);

// Make sure you add code to validate these fields
var luisAppId = "f9f21b33-1436-4de9-ba02-5fe838321101";
var luisAPIKey = "6539064cc46f4acb902501c0fdba5ffd";
var luisAPIHostName = process.env.LuisAPIHostName || 'westus.api.cognitive.microsoft.com';

const LuisModelUrl = 'https://' + luisAPIHostName + '/luis/v2.0/apps/' + luisAppId + '?subscription-key=' + luisAPIKey;


// Create a recognizer that gets intents from LUIS, and add it to the bot
var recognizer = new builder.LuisRecognizer(LuisModelUrl);
var intents = new builder.IntentDialog({
  recognizers: [recognizer]
});

bot.dialog("/", intents);


//Dialog for weather related questions
bot.dialog("/weather", function (session) {
    session.send("It's always a bright sunny day for me. Hope it's the same for you. I am here to help you with your invoice and payment related queries. Type \"help\" or \"menu\" for options or continue asking a question");
    session.endDialog();
//}).triggerAction({ matches: /^(weather|season|How is the weather today)/i })
}).triggerAction({matches : /^(^.*\b(weather|season|hot|cold|rain|humid)\b.*$)/i})

//Dialog for saying sorry as user is not happy with help
bot.dialog("/nothelping", function (session) {
    session.send("I am sorry I wasn't able to help you. Let me to connect you to a representative who can help further !");
    session.endDialog();
//}).triggerAction({ matches: /^(weather|season|How is the weather today)/i })
}).triggerAction({matches : /^(^.*\b(This is not helping|not helping|Unable to get proper information)\b.*$)/i})

//Dialog for one FAQ asking how to submit an invoice
bot.dialog("/howtosubmitinvoice", function (session) {
    session.send("Your account needs to be configured to handle invoices through your customer's Coupa instance. Contact them to get set up with an invoicing method.");
    session.send("Visit https://www.coupa.com/ for more information")
    session.endDialog();
//}).triggerAction({ matches: /^(weather|season|How is the weather today)/i })
}).triggerAction({matches : /^(^.*\b(How do I submit an invoice?|How to submit an invoice ?|How can I submit an invoice ?)\b.*$)/i})


//Dialog for FAQs
bot.dialog("/faqs", [ 
  function (session, args, next) {
    builder.Prompts.choice(session, "Here are some Frequently Asked Questions..", "How do I submit an invoice? | Can I reuse an invoice number once my customer has deleted it on their end? | How do I know if an invoice has been registered? |How do I add shipping charges to an invoice?", {listStyle:builder.ListStyle.button });

  },
  function (session, results) {
    switch (results.response.index) {
      case 0:
          session.send("Your account needs to be configured to handle invoices through your customer's Coupa instance. Contact them to get set up with an invoicing method.")
          session.endDialog();
          break;
      case 1:
          session.send("Yes, if a customer deleted an invoice number, you can reuse it.")
          session.endDialog();
          break;
      case 2:
          session.send("On the main menu, click on the Invoices tab. For the desired customer, look for invoices that are still listed as drafts. They need to be edited and then submitted. If the invoice is not there, it's not in the system.")
          session.endDialog();
          break;
      case 3: 
          session.send("You can add shipping charges at the bottom of the invoice or at the line item level. Ask your customer which method they prefer.")
        session.endDialog();
          break;
      default:
          session.endDialog();
          break;
  } //end switch
  }
]).triggerAction({matches : /^(^.*\b(FAQ|FAQs|Show me frequently asked questions|frequently asked questions)\b.*$)/i})

//Intent for Payment Status 
intents.matches("PaymentStatus", [
  function (session, args, next) {
    session.send("It generally takes 30 days from the date of application. I would need some information to get exact details.");
    next();
  }, 
  function (session, args, next){
    session.beginDialog("askInvoiceID");
    console.log(session.userData.invoiceID);
  },
  function (session, args, next) {
    var invoiceid = session.userData.invoiceID;
      var url = "http://40.121.88.94:8081/invoice/" + invoiceid.toUpperCase()
      request.get(url, (error, response, body) => {
      if(error) {
          return console.dir(error);
      } //end if 
      var data = JSON.parse(body);
      if (data.length == 0) {
          session.send("Sorry ! Seems like there is no invoice on this given Invoice ID");
      } else { 
      //Get the due date for invoice   
      var due_date_string = data[0]['invoice_due_date']
      var __due_date = stringToDate(due_date_string,"mm-dd-yyyy","-")
      //Calculate todays date
      var today = new Date();
      var dd = today.getDate();
      var mm = today.getMonth() + 1; //January is 0!
      var yyyy = today.getFullYear();
      date_today_string = mm + "-" + dd + "-" + yyyy
      var __today = stringToDate(date_today_string, "mm-dd-yyyy", "-")
      
      //console.log("Today " + date_today_string);
      //console.log("Due Date - " + due_date_string);

      var timeDiff  = Math.abs(__due_date.getTime() - __today.getTime())
      var diffDays = Math.ceil(timeDiff / (1000 * 3600 * 24));
      if(__today > __due_date) {
          session.send("Ohh ! I see the due date for payment is elapsed by " + diffDays + " days.")
          session.send("I will treat this as urgent and regret for the inconvinience. Transaferring call to our executive to dig deeper into the issue. ")
      } else {
          session.send("Hi ! I see that the due date for payment is  " + due_date_string + " which has " + diffDays + " days to come." );
          session.send("We request you to please wait till " + due_date_string + " for payment to be released.")
          session.send("In case of any other help, please type \"help\" or \"menu\"");
      }
      } //end else


  })//end request
} //end first function  
]);

//Intent for Greetings
intents.matches('Greeting', function (session, args, next) {
  session.send("Hello ! I am Eva. I can assist you with your invoice and payment related queries !");
  session.send("Type \"help\" or \"menu\" to get the list of items I can help you with or ask a question related to your invoice and payments.");
}) ;

//Intent for Help
intents.matches('Help', [
  function (session, args, next) {
  builder.Prompts.choice(session, "I can help you on the following !", "Get Pending invoice information | Get Approved invoice information | Get Pending payments information | Get Approved payments information ", { listStyle: builder.ListStyle.button });
}, 
function (session, results) {
  console.log(results.response.entity);
  switch (results.response.index) {
      case 0:
          session.beginDialog('allPendingInvoice');
          break;
      case 1:
          session.beginDialog("allApprovedInvoice");
          break;
      case 2:
          session.beginDialog("allPendingPayments");
          break;
      case 3: 
          session.beginDialog("allApprovedPayments");
          break;
      default:
          session.endDialog();
          break;
  } //end switch
}
]);

//Intent for supplier status
intents.matches("SupplierStatus", [
  function (session, args, next) {
    //console.log("Inside supplier intent")
    var supplierEntity = builder.EntityRecognizer.findEntity(args.entities, "SupplierID");
    if(supplierEntity == null) {
      delete session.userData.supplierID;
    } else {
      session.userData.supplierID = supplierEntity.entity;
    }
    if (!session.userData.supplierID) {
      session.beginDialog("askSupplierID")
    } else {
      next();
    }
  }, 
  function (session, args, next) {
    var supplierid = session.userData.supplierID;
      var url = "http://40.121.88.94:8081/supplier/" + supplierid.toUpperCase()
      request.get(url, (error, response, body) => {
      if(error) {
          return console.dir(error);
      } //end if 
      var data = JSON.parse(body);
      if (data.length == 0) {
          session.send("Sorry ! Seems like there is no Supplier on this given Supplier ID");
      } else { 
        
      for (var i = 0; i < data.length; i++) {
          session.send("Hey ! I found the following records -");
          var __invoice_id = data[i]['invoice_id'];
          var __invoice_status = data[i]['invoice_status'];
          var __balance_due = data[i]['balance_due'];
          var __invoice_due_date = data[i]['invoice_due_date'];
          var __approver =  data[i]['approver'];
          var __supplier_name = data[i]['supplier_name'];
          var __payment_status = data[i]['payment_status'];
          var __invoice_date = data[i]['invoice_date'];


          var card_json = {
              "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
              "type": "AdaptiveCard",
              "version": "1.0",
              "speak": "",
              "body": [
                {
                  "type": "ColumnSet",
                  "columns": [
                    {
                      "type": "Column",
                      "width": "auto",
                      "items": [
                       {
                          "type": "TextBlock",
                          "text": __invoice_id,
                          "isSubtle": true,
                          "size": "large",
                          "weight": "bolder"
                        }
                      ]
                    },
                    {
                      "type": "Column",
                      "width": "stretch",
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": "Invoice Status",
                          "horizontalAlignment": "right",
                          "isSubtle": true
                        },
                        {
                          "type": "TextBlock",
                          "text": __invoice_status,
                          "horizontalAlignment": "right",
                          "spacing": "none",
                          "size": "large",
                          "color": "attention"
                        }
                      ]
                    }
                  ]
                },
                {
                  "type": "ColumnSet",
                  "separator": true,
                  "spacing": "medium",
                  "columns": [
                    {
                      "type": "Column",
                      "width": "stretch",
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": " ",
                          "isSubtle": true,
                          "weight": "bolder"
                        },
                        {
                          "type": "TextBlock",
                          "text": __supplier_name,
                          "spacing": "small"
                        },
                        {
                          "type": "TextBlock",
                          "text": __approver,
                          "spacing": "small"
                        }
                      ]
                    },
                    {
                      "type": "Column",
                      "width": "auto",
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": " ",
                          "horizontalAlignment": "right",
                          "isSubtle": true,
                          "weight": "bolder"
                        },
                        {
                          "type": "TextBlock",
                          "text": "Supplier",
                          "horizontalAlignment": "right",
                          "spacing": "small"
                        },
                        {
                          "type": "TextBlock",
                          "text": "Approver",
                          "horizontalAlignment": "right",
                          "spacing": "small"
                        }
                      ]
                    }
                  ]
                },
                {
                  "type": "ColumnSet",
                  "spacing": "medium",
                  "separator": true,
                  "columns": [
                    {
                      "type": "Column",
                      "width": 1,
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": "Balance Due",
                          "isSubtle": true,
                          "weight": "bolder"
                        },
                        {
                          "type": "TextBlock",
                          "text": __balance_due,
                          "spacing": "small"
                        }
                      ]
                    },
                    {
                      "type": "Column",
                      "width": 1,
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": "Payment Status",
                          "isSubtle": true,
                          "weight": "bolder"
                        },
                        {
                          "type": "TextBlock",
                          "text": __payment_status,
                          "spacing": "small"
                        }
                      ]
                    },
                    {
                      "type": "Column",
                      "width": 1,
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": "Invoice Submitted Date",
                          "isSubtle": true,
                          "horizontalAlignment": "right",
                          "weight": "bolder"
                        },
                        {
                          "type": "TextBlock",
                          "text": __invoice_date,
                          "color": "attention",
                          "horizontalAlignment": "right",
                          "weight": "bolder",
                          "spacing": "small"
                        }
                      ]
                    }
                  ]
                }
              ]
            } //end of card
      var adaptiveCardMessage = new builder.Message(session)
      .addAttachment({
      contentType: "application/vnd.microsoft.card.adaptive",
      content: card_json
      });session.send(adaptiveCardMessage);
      } //end for loop   
  }//end else
     
    }) ; //end request    
    session.endDialog();
} //end function
  
])

//Intent for invoice Status 
intents.matches("InvoiceStatus", [
  function (session, args, next) {
    //session.send("Inside invoice status intent");
    //console.log(JSON.stringify(args));
    //console.log(args.intent);
  
    var invoiceEntity = builder.EntityRecognizer.findEntity(args.entities, "InvoiceID");
    if (invoiceEntity == null) {
        delete session.userData.invoiceID;
    } else {
        session.userData.invoiceID = invoiceEntity.entity; 
    }   
    if (!session.userData.invoiceID) {
        session.beginDialog("askInvoiceID");
    } else {
       next();
    }
  }, 
    function (session, args, next) {
      //console.log("OK");
      //console.log(session.userData.invoiceID)
      //Call Request function here 
      var invoiceid = session.userData.invoiceID;
      var url = "http://40.121.88.94:8081/invoice/" + invoiceid.toUpperCase()
      request.get(url, (error, response, body) => {
      if(error) {
          return console.dir(error);
      } //end if 
      var data = JSON.parse(body);
      if (data.length == 0) {
          session.send("Sorry ! Seems like there is no invoice on this given Invoice ID");
      } else { 
      session.send("Hey ! I found the following records -");  
      for (var i = 0; i < data.length; i++) {
    
          var __invoice_id = data[i]['invoice_id'];
          var __invoice_status = data[i]['invoice_status'];
          var __balance_due = data[i]['balance_due'];
          var __invoice_due_date = data[i]['invoice_due_date'];
          var __approver =  data[i]['approver'];
          var __supplier_name = data[i]['supplier_name'];
          var __payment_status = data[i]['payment_status'];
          var __invoice_date = data[i]['invoice_date'];
          var __color_to_use = "";
          if (__invoice_status == "Pending") {
              __color_to_use = "warning" ;
          } else {
              __color_to_use = "good";
          }


          var card_json = {
              "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
              "type": "AdaptiveCard",
              "version": "1.0",
              "speak": "",
              "body": [
                {
                  "type": "ColumnSet",
                  "columns": [
                    {
                      "type": "Column",
                      "width": "auto",
                      "items": [
                       {
                          "type": "TextBlock",
                          "text": __invoice_id,
                          "isSubtle": true,
                          "size": "large",
                          "weight": "bolder"
                        }
                      ]
                    },
                    {
                      "type": "Column",
                      "width": "stretch",
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": "Invoice Status",
                          "horizontalAlignment": "right",
                          "isSubtle": true
                        },
                        {
                          "type": "TextBlock",
                          "text": __invoice_status,
                          "horizontalAlignment": "right",
                          "spacing": "none",
                          "size": "large",
                          "color": __color_to_use
                        }
                      ]
                    }
                  ]
                },
                {
                  "type": "ColumnSet",
                  "separator": true,
                  "spacing": "medium",
                  "columns": [
                    {
                      "type": "Column",
                      "width": "stretch",
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": " ",
                          "isSubtle": true,
                          "weight": "bolder"
                        },
                        {
                          "type": "TextBlock",
                          "text": __supplier_name,
                          "spacing": "small"
                        },
                        {
                          "type": "TextBlock",
                          "text": __approver,
                          "spacing": "small"
                        }
                      ]
                    },
                    {
                      "type": "Column",
                      "width": "auto",
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": " ",
                          "horizontalAlignment": "right",
                          "isSubtle": true,
                          "weight": "bolder"
                        },
                        {
                          "type": "TextBlock",
                          "text": "Supplier",
                          "horizontalAlignment": "right",
                          "spacing": "small"
                        },
                        {
                          "type": "TextBlock",
                          "text": "Approver",
                          "horizontalAlignment": "right",
                          "spacing": "small"
                        }
                      ]
                    }
                  ]
                },
                {
                  "type": "ColumnSet",
                  "spacing": "medium",
                  "separator": true,
                  "columns": [
                    {
                      "type": "Column",
                      "width": 1,
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": "Balance Due",
                          "isSubtle": true,
                          "weight": "bolder"
                        },
                        {
                          "type": "TextBlock",
                          "text": __balance_due,
                          "spacing": "small"
                        }
                      ]
                    },
                    {
                      "type": "Column",
                      "width": 1,
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": "Payment Status",
                          "isSubtle": true,
                          "weight": "bolder"
                        },
                        {
                          "type": "TextBlock",
                          "text": __payment_status,
                          "spacing": "small"
                        }
                      ]
                    },
                    {
                      "type": "Column",
                      "width": 1,
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": "Invoice Submitted Date",
                          "isSubtle": true,
                          "horizontalAlignment": "right",
                          "weight": "bolder"
                        },
                        {
                          "type": "TextBlock",
                          "text": __invoice_date,
                          "horizontalAlignment": "right",
                          "spacing": "small"
                        }
                      ]
                    }
                  ]
                }
              ]
            } //end of card
      var adaptiveCardMessage = new builder.Message(session)
      .addAttachment({
      contentType: "application/vnd.microsoft.card.adaptive",
      content: card_json
      });session.send(adaptiveCardMessage);
      } //end for loop   
  }//end else
}) ; //end request    
    session.endDialog();
} //end function
]);
 
//Dialog for requesting invoice ID
bot.dialog('askInvoiceID', [
  function (session, args, next) {
  builder.Prompts.text(session, "Please share your invoice ID ...");
}, function (session, results) {
  session.userData.invoiceID = results.response;
  session.endDialogWithResult(results);
}
]);

//Dialog for all pending invoice 
bot.dialog("allPendingInvoice", [
  function (session, args, next) {
     session.beginDialog('askSupplierID')
    /*if (session.userData.supplier_id) {
     
    } else {
      next()
    }*/
  }, 
  function (session, results) {
    //console.log("Supplier name is " + session.userData.supplier_name);
    var supplier_id = session.userData.supplier_id;
    var url = "http://40.121.88.94:8081/supplier/pending/" + supplier_id.toUpperCase();
      request.get(url, (error, response, body) => {
      if(error) {
          return console.dir(error);
      } //end if 
      var data = JSON.parse(body);
      console.log(data);
      if (data.length == 0) {
          session.send("Sorry ! Seems like there is no Pending invoice on this given Supplier ID");
          session.endDialog();
      } else { 
      session.send("Hey ! I found the following records -"); 
      for (var i = 0; i < data.length; i++) {
        
        if(data[i]['invoice_status'] == "Pending") {
          
          var __invoice_id = data[i]['invoice_id'];
          var __invoice_status = data[i]['invoice_status'];
          var __balance_due = data[i]['balance_due'];
          var __invoice_due_date = data[i]['invoice_due_date'];
          var __approver =  data[i]['approver'];
          var __supplier_name = data[i]['supplier_name'];
          var __payment_status = data[i]['payment_status'];
          var __invoice_date = data[i]['invoice_date'];


          var card_json = {
              "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
              "type": "AdaptiveCard",
              "version": "1.0",
              "speak": "",
              "body": [
                {
                  "type": "ColumnSet",
                  "columns": [
                    {
                      "type": "Column",
                      "width": "auto",
                      "items": [
                         {
                          "type": "TextBlock",
                          "text": __invoice_id,
                          "isSubtle": true,
                          "size": "large",
                          "weight": "bolder"
                        }
                      ]
                    },
                    {
                      "type": "Column",
                      "width": "stretch",
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": "Invoice Status",
                          "horizontalAlignment": "right",
                          "isSubtle": true
                        },
                        {
                          "type": "TextBlock",
                          "text": __invoice_status,
                          "horizontalAlignment": "right",
                          "spacing": "none",
                          "size": "large",
                          "color": "warning"
                        }
                      ]
                    }
                  ]
                },
                {
                  "type": "ColumnSet",
                  "separator": true,
                  "spacing": "medium",
                  "columns": [
                    {
                      "type": "Column",
                      "width": "stretch",
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": " ",
                          "isSubtle": true,
                          "weight": "bolder"
                        },
                        {
                          "type": "TextBlock",
                          "text": __supplier_name,
                          "spacing": "small"
                        },
                        {
                          "type": "TextBlock",
                          "text": __approver,
                          "spacing": "small"
                        }
                      ]
                    },
                    {
                      "type": "Column",
                      "width": "auto",
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": " ",
                          "horizontalAlignment": "right",
                          "isSubtle": true,
                          "weight": "bolder"
                        },
                        {
                          "type": "TextBlock",
                          "text": "Supplier",
                          "horizontalAlignment": "right",
                          "spacing": "small"
                        },
                        {
                          "type": "TextBlock",
                          "text": "Approver",
                          "horizontalAlignment": "right",
                          "spacing": "small"
                        }
                      ]
                    }
                  ]
                },
                {
                  "type": "ColumnSet",
                  "spacing": "medium",
                  "separator": true,
                  "columns": [
                    {
                      "type": "Column",
                      "width": 1,
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": "Balance Due",
                          "isSubtle": true,
                          "weight": "bolder"
                        },
                        {
                          "type": "TextBlock",
                          "text": __balance_due,
                          "spacing": "small"
                        }
                      ]
                    },
                    {
                      "type": "Column",
                      "width": 1,
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": "Payment Status",
                          "isSubtle": true,
                          "weight": "bolder"
                        },
                        {
                          "type": "TextBlock",
                          "text": __payment_status,
                          "spacing": "small"
                        }
                      ]
                    },
                    {
                      "type": "Column",
                      "width": 1,
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": "Invoice Submitted Date",
                          "isSubtle": true,
                          "horizontalAlignment": "right",
                          "weight": "bolder"
                        },
                        {
                          "type": "TextBlock",
                          "text": __invoice_date,
                          "horizontalAlignment": "right",
                          "spacing": "small"
                        }
                      ]
                    }
                  ]
                }
              ]
            } //end of card
      var adaptiveCardMessage = new builder.Message(session)
      .addAttachment({
      contentType: "application/vnd.microsoft.card.adaptive",
      content: card_json
      });session.send(adaptiveCardMessage);
       } //end if 
      } //end for loop   
  }//end else     
    }) //end of request
session.endDialog();
  }
]);

//Dialog for all Approved invoice 
bot.dialog("allApprovedInvoice", [
  function (session, args, next) {
    session.beginDialog('askSupplierID')
    /*if (session.userData.supplier_id) {
      session.beginDialog('askSupplierID')
    } else {
      next()
    }*/
  }, 
  function (session, results) {
    //console.log("Supplier name is " + session.userData.supplier_name);
    var supplier_id = session.userData.supplier_id;
    var url = "http://40.121.88.94:8081/supplier/approved/" + supplier_id.toUpperCase();
      request.get(url, (error, response, body) => {
      if(error) {
          return console.dir(error);
      } //end if 
      var data = JSON.parse(body);
      if (data.length == 0) {
          session.send("Sorry ! Seems like there is no invoice on this given Supplier ID with Approved status");
          session.endDialog();
      } else { 
      session.send("Hey ! I found the following records -");  
      for (var i = 0; i < data.length; i++) {
        if(data[i]['invoice_status'] == "Approved") {
          var __invoice_id = data[i]['invoice_id'];
          var __invoice_status = data[i]['invoice_status'];
          var __balance_due = data[i]['balance_due'];
          var __invoice_due_date = data[i]['invoice_due_date'];
          var __approver =  data[i]['approver'];
          var __supplier_name = data[i]['supplier_name'];
          var __payment_status = data[i]['payment_status'];
          var __invoice_date = data[i]['invoice_date'];


          var card_json = {
              "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
              "type": "AdaptiveCard",
              "version": "1.0",
              "speak": "",
              "body": [
                {
                  "type": "ColumnSet",
                  "columns": [
                    {
                      "type": "Column",
                      "width": "auto",
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": __invoice_id,
                          "isSubtle": true,
                          "size": "large",
                          "weight": "bolder"
                        }
                      ]
                    },
                    {
                      "type": "Column",
                      "width": "stretch",
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": "Invoice Status",
                          "horizontalAlignment": "right",
                          "isSubtle": true
                        },
                        {
                          "type": "TextBlock",
                          "text": __invoice_status,
                          "horizontalAlignment": "right",
                          "spacing": "none",
                          "size": "large",
                          "color": "good"
                        }
                      ]
                    }
                  ]
                },
                {
                  "type": "ColumnSet",
                  "separator": true,
                  "spacing": "medium",
                  "columns": [
                    {
                      "type": "Column",
                      "width": "stretch",
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": " ",
                          "isSubtle": true,
                          "weight": "bolder"
                        },
                        {
                          "type": "TextBlock",
                          "text": __supplier_name,
                          "spacing": "small"
                        },
                        {
                          "type": "TextBlock",
                          "text": __approver,
                          "spacing": "small"
                        }
                      ]
                    },
                    {
                      "type": "Column",
                      "width": "auto",
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": " ",
                          "horizontalAlignment": "right",
                          "isSubtle": true,
                          "weight": "bolder"
                        },
                        {
                          "type": "TextBlock",
                          "text": "Supplier",
                          "horizontalAlignment": "right",
                          "spacing": "small"
                        },
                        {
                          "type": "TextBlock",
                          "text": "Approver",
                          "horizontalAlignment": "right",
                          "spacing": "small"
                        }
                      ]
                    }
                  ]
                },
                {
                  "type": "ColumnSet",
                  "spacing": "medium",
                  "separator": true,
                  "columns": [
                    {
                      "type": "Column",
                      "width": 1,
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": "Balance Due",
                          "isSubtle": true,
                          "weight": "bolder"
                        },
                        {
                          "type": "TextBlock",
                          "text": __balance_due,
                          "spacing": "small"
                        }
                      ]
                    },
                    {
                      "type": "Column",
                      "width": 1,
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": "Payment Status",
                          "isSubtle": true,
                          "weight": "bolder"
                        },
                        {
                          "type": "TextBlock",
                          "text": __payment_status,
                          "spacing": "small"
                        }
                      ]
                    },
                    {
                      "type": "Column",
                      "width": 1,
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": "Invoice Submitted Date",
                          "isSubtle": true,
                          "horizontalAlignment": "right",
                          "weight": "bolder"
                        },
                        {
                          "type": "TextBlock",
                          "text": __invoice_date,
                          "horizontalAlignment": "right",
                          "spacing": "small"
                        }
                      ]
                    }
                  ]
                }
              ]
            } //end of card
      var adaptiveCardMessage = new builder.Message(session)
      .addAttachment({
      contentType: "application/vnd.microsoft.card.adaptive",
      content: card_json
      });session.send(adaptiveCardMessage);
       } //end of if
      } //end for loop   
  }//end else
     
    }) //end of request
session.endDialog();
  }

]);

//Dialog for all pending payments
bot.dialog("allPendingPayments", [
  function (session, args, next) {
    session.beginDialog('askSupplierID')
    /*if (session.userData.supplier_id) {
      session.beginDialog('askSupplierID')
    } else {
      next()
    }*/
  }, 
  function (session, results) {
    //console.log("Supplier name is " + session.userData.supplier_name);
    var supplier_id = session.userData.supplier_id;
    var url = "http://40.121.88.94:8081/supplier/" + supplier_id.toUpperCase();
      request.get(url, (error, response, body) => {
      if(error) {
          return console.dir(error);
      } //end if 
      var data = JSON.parse(body);
      if (data.length == 0) {
          session.send("Sorry ! Seems like there is no invoice on this given Supplier ID");
          session.endDialog();
      } else { 
       
      var found = 0
      for (var i = 0; i < data.length; i++) {
        if(data[i]['payment_status'] == "Pending") {  
          found = 1;
          
          var __invoice_id = data[i]['invoice_id'];
          var __invoice_status = data[i]['invoice_status'];
          var __balance_due = data[i]['balance_due'];
          var __invoice_due_date = data[i]['invoice_due_date'];
          var __approver =  data[i]['approver'];
          var __supplier_name = data[i]['supplier_name'];
          var __payment_status = data[i]['payment_status'];
          var __invoice_date = data[i]['invoice_date'];
          var __color_to_use = "" ;
          if (__invoice_status == "Pending") {
            __color_to_use = "warning";
          } else {
            __color_to_use = "good";
          }

          var card_json = {
              "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
              "type": "AdaptiveCard",
              "version": "1.0",
              "speak": "",
              "body": [
                {
                  "type": "ColumnSet",
                  "columns": [
                    {
                      "type": "Column",
                      "width": "auto",
                      "items": [
                       {
                          "type": "TextBlock",
                          "text": __invoice_id,
                          "isSubtle": true,
                          "size": "large",
                          "weight": "bolder"
                        }
                      ]
                    },
                    {
                      "type": "Column",
                      "width": "stretch",
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": "Invoice Status",
                          "horizontalAlignment": "right",
                          "isSubtle": true
                        },
                        {
                          "type": "TextBlock",
                          "text": __invoice_status,
                          "horizontalAlignment": "right",
                          "spacing": "none",
                          "size": "large",
                          "color": __color_to_use
                        }
                      ]
                    }
                  ]
                },
                {
                  "type": "ColumnSet",
                  "separator": true,
                  "spacing": "medium",
                  "columns": [
                    {
                      "type": "Column",
                      "width": "stretch",
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": " ",
                          "isSubtle": true,
                          "weight": "bolder"
                        },
                        {
                          "type": "TextBlock",
                          "text": __supplier_name,
                          "spacing": "small"
                        },
                        {
                          "type": "TextBlock",
                          "text": __approver,
                          "spacing": "small"
                        }
                      ]
                    },
                    {
                      "type": "Column",
                      "width": "auto",
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": " ",
                          "horizontalAlignment": "right",
                          "isSubtle": true,
                          "weight": "bolder"
                        },
                        {
                          "type": "TextBlock",
                          "text": "Supplier",
                          "horizontalAlignment": "right",
                          "spacing": "small"
                        },
                        {
                          "type": "TextBlock",
                          "text": "Approver",
                          "horizontalAlignment": "right",
                          "spacing": "small"
                        }
                      ]
                    }
                  ]
                },
                {
                  "type": "ColumnSet",
                  "spacing": "medium",
                  "separator": true,
                  "columns": [
                    {
                      "type": "Column",
                      "width": 1,
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": "Balance Due",
                          "isSubtle": true,
                          "weight": "bolder"
                        },
                        {
                          "type": "TextBlock",
                          "text": __balance_due,
                          "spacing": "small"
                        }
                      ]
                    },
                    {
                      "type": "Column",
                      "width": 1,
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": "Payment Status",
                          "isSubtle": true,
                          "weight": "bolder"
                        },
                        {
                          "type": "TextBlock",
                          "text": __payment_status,
                          "spacing": "small",
                          "color": "warning"
                        }
                      ]
                    },
                    {
                      "type": "Column",
                      "width": 1,
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": "Invoice Submitted Date",
                          "isSubtle": true,
                          "horizontalAlignment": "right",
                          "weight": "bolder"
                        },
                        {
                          "type": "TextBlock",
                          "text": __invoice_date,
                          "horizontalAlignment": "right",
                          "spacing": "small"
                        }
                      ]
                    }
                  ]
                }
              ]
            } //end of card
      var adaptiveCardMessage = new builder.Message(session)
      .addAttachment({
      contentType: "application/vnd.microsoft.card.adaptive",
      content: card_json
      });session.send(adaptiveCardMessage);
    } //end if  
    } //end for loop   
    if (found == 0) {
      session.send("Seems like there are no Pending Payments on this given Supplier ID.")
    }
  }//end else
     
    }) //end of request
session.endDialog();
  }

]);

//Dialog for all pending payments
bot.dialog("allApprovedPayments", [
  function (session, args, next) {
     session.beginDialog('askSupplierID')
    /*if (session.userData.supplier_id) {
      session.beginDialog('askSupplierID')
    } else {
      next()
    }*/
  }, 
  function (session, results) {
    //console.log("Supplier name is " + session.userData.supplier_name);
    var supplier_id = session.userData.supplier_id;
    var url = "http://40.121.88.94:8081/supplier/" + supplier_id.toUpperCase();
      request.get(url, (error, response, body) => {
      if(error) {
          return console.dir(error);
      } //end if 
      var data = JSON.parse(body);
      if (data.length == 0) {
          session.send("Sorry ! Seems like there is no invoice on this given Supplier ID");
          session.endDialog();
      } else { 
       var found = 0; //init found to 0 if remains 0 means no record was found
      for (var i = 0; i < data.length; i++) {
        if(data[i]['payment_status'] == "Approved") {  
          found = 1; //means a record was found
          
          var __invoice_id = data[i]['invoice_id'];
          var __invoice_status = data[i]['invoice_status'];
          var __balance_due = data[i]['balance_due'];
          var __invoice_due_date = data[i]['invoice_due_date'];
          var __approver =  data[i]['approver'];
          var __supplier_name = data[i]['supplier_name'];
          var __payment_status = data[i]['payment_status'];
          var __invoice_date = data[i]['invoice_date'];


          var card_json = {
              "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
              "type": "AdaptiveCard",
              "version": "1.0",
              "speak": "",
              "body": [
                {
                  "type": "ColumnSet",
                  "columns": [
                    {
                      "type": "Column",
                      "width": "auto",
                      "items": [
                       {
                          "type": "TextBlock",
                          "text": __invoice_id,
                          "isSubtle": true,
                          "size": "large",
                          "weight": "bolder"
                        }
                      ]
                    },
                    {
                      "type": "Column",
                      "width": "stretch",
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": "Invoice Status",
                          "horizontalAlignment": "right",
                          "isSubtle": true
                        },
                        {
                          "type": "TextBlock",
                          "text": __invoice_status,
                          "horizontalAlignment": "right",
                          "spacing": "none",
                          "size": "large",
                          "color": "good"
                        }
                      ]
                    }
                  ]
                },
                {
                  "type": "ColumnSet",
                  "separator": true,
                  "spacing": "medium",
                  "columns": [
                    {
                      "type": "Column",
                      "width": "stretch",
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": " ",
                          "isSubtle": true,
                          "weight": "bolder"
                        },
                        {
                          "type": "TextBlock",
                          "text": __supplier_name,
                          "spacing": "small"
                        },
                        {
                          "type": "TextBlock",
                          "text": __approver,
                          "spacing": "small"
                        }
                      ]
                    },
                    {
                      "type": "Column",
                      "width": "auto",
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": " ",
                          "horizontalAlignment": "right",
                          "isSubtle": true,
                          "weight": "bolder"
                        },
                        {
                          "type": "TextBlock",
                          "text": "Supplier",
                          "horizontalAlignment": "right",
                          "spacing": "small"
                        },
                        {
                          "type": "TextBlock",
                          "text": "Approver",
                          "horizontalAlignment": "right",
                          "spacing": "small"
                        }
                      ]
                    }
                  ]
                },
                {
                  "type": "ColumnSet",
                  "spacing": "medium",
                  "separator": true,
                  "columns": [
                    {
                      "type": "Column",
                      "width": 1,
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": "Balance Due",
                          "isSubtle": true,
                          "weight": "bolder"
                        },
                        {
                          "type": "TextBlock",
                          "text": __balance_due,
                          "spacing": "small"
                        }
                      ]
                    },
                    {
                      "type": "Column",
                      "width": 1,
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": "Payment Status",
                          "isSubtle": true,
                          "weight": "bolder"
                        },
                        {
                          "type": "TextBlock",
                          "text": __payment_status,
                          "spacing": "small"
                        }
                      ]
                    },
                    {
                      "type": "Column",
                      "width": 1,
                      "items": [
                        {
                          "type": "TextBlock",
                          "text": "Invoice Submitted Date",
                          "isSubtle": true,
                          "horizontalAlignment": "right",
                          "weight": "bolder"
                        },
                        {
                          "type": "TextBlock",
                          "text": __invoice_date,
                          "color": "attention",
                          "horizontalAlignment": "right",
                          "weight": "bolder",
                          "spacing": "small"
                        }
                      ]
                    }
                  ]
                }
              ]
            } //end of card
      var adaptiveCardMessage = new builder.Message(session)
      .addAttachment({
      contentType: "application/vnd.microsoft.card.adaptive",
      content: card_json
      });session.send(adaptiveCardMessage);
    } //end if  
    } //end for loop   
    if (found == 0) {
      session.send("Seems like there are no Approved Payments on this given Supplier ID")
    }
  }//end else
     
    }) //end of request
session.endDialog();
  }

]);



//Dialog to ask for supplier ID
bot.dialog("askSupplierID", [
  function (session, args, next) {
    builder.Prompts.text(session, "Please provide your supplier ID ");
  }, 
  function (session, results) {
    session.userData.supplier_id = results.response
    session.endDialogWithResult(results);
  }
]);


//function to convert STring format date to Date format 
function stringToDate(_date,_format,_delimiter)
{
            var formatLowerCase=_format.toLowerCase();
            var formatItems=formatLowerCase.split(_delimiter);
            var dateItems=_date.split(_delimiter);
            var monthIndex=formatItems.indexOf("mm");
            var dayIndex=formatItems.indexOf("dd");
            var yearIndex=formatItems.indexOf("yyyy");
            var month=parseInt(dateItems[monthIndex]);
            month-=1;
            var formatedDate = new Date(dateItems[yearIndex],month,dateItems[dayIndex]);
            return formatedDate;
}
