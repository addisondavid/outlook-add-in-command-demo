// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.

/// <reference path="../App.js" />

  "use strict";

  class MailItem {
    constructor(item) {
      this.item = item;
    }
    get(prop) {
      return new Promise((resolve, reject) => {
        if (!this.item[prop] || !this.item[prop].getAsync) return resolve(this.item[prop]);
        this.item[prop].getAsync(res => {
          resolve(res && res.value);
        });
      });
    }
  }

  // The Office initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
      console.log('debugfilter', 'office.initialize');
      $(document).ready(function () {
          console.log('debugfilter', 'document.ready');
          app.initialize();

          console.log('debugfilter', 'isPersistenceSupported', isPersistenceSupported());
          if (isPersistenceSupported()) {
            // Set up ItemChanged event
            // Office.EventType.ItemChanged === 'olkItemSelectedChanged'
            Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem);
          }

          console.log('debugfilter', 'loadProps');
          loadProps(Office.context.mailbox.item);
          $('#action-button').click(openDialog);
          $('#action-button2').click(openDialogAsIframe);
      });
  };

  function isPersistenceSupported() {
    // This feature is part of the preview 1.5 req set
    // Since 1.5 isn't fully implemented, just check that the
    // method is defined.
    // Once 1.5 is implemented, we can replace this with
    // Office.context.requirements.isSetSupported('Mailbox', 1.5)
    return Office.context.mailbox.addHandlerAsync !== undefined;
  };

  function loadNewItem(eventArgs) {
    loadProps(Office.context.mailbox.item);
  };

  // Take an array of AttachmentDetails objects and
  // build a list of attachment names, separated by a line-break
  function buildAttachmentsString(attachments) {
    if (attachments && attachments.length > 0) {
      var returnString = "";

      for (var i = 0; i < attachments.length; i++) {
        if (i > 0) {
          returnString = returnString + "<br/>";
        }
        returnString = returnString + attachments[i].name;
      }

      return returnString;
    }

    return "None";
  }

  // Format an EmailAddressDetails object as
  // GivenName Surname <emailaddress>
  function buildEmailAddressString(address) {
    return address && address.displayName + " &lt;" + address.emailAddress + "&gt;";
  }

  // Take an array of EmailAddressDetails objects and
  // build a list of formatted strings, separated by a line-break
  function buildEmailAddressesString(addresses) {
    if (addresses && addresses.length > 0) {
      var returnString = "";

      for (var i = 0; i < addresses.length; i++) {
        if (i > 0) {
          returnString = returnString + "<br/>";
        }
        returnString = returnString + buildEmailAddressString(addresses[i]);
      }

      return returnString;
    }

    return "None";
  }

  // Load properties from a Message object
  function loadMessageProps(item) {
    const mailItem = new MailItem(item);
    $('#message-props').show();

    mailItem.get('attachments').then(attachments => $('#attachments').html(buildAttachmentsString(attachments)));
    mailItem.get('cc').then(cc => $('#cc').html(buildEmailAddressesString(cc)));
    mailItem.get('conversationId').then(conversationId => $('#conversationId').text(conversationId));
    mailItem.get('from').then(from => $('#from').html(buildEmailAddressString(from)));
    mailItem.get('internetMessageId').then(internetMessageId => $('#internetMessageId').text(internetMessageId));
    mailItem.get('normalizedSubject').then(normalizedSubject => $('#normalizedSubject').text(normalizedSubject));
    mailItem.get('sender').then(sender => $('#sender').html(buildEmailAddressString(sender)));
    mailItem.get('subject').then(subject => $('#subject').text(subject));
    mailItem.get('to').then(to => $('#to').html(buildEmailAddressesString(to)));
  }

  // Load properties from the Item base object, then load the
  // type-specific properties.
  function loadProps(item) {

    $('#dateTimeCreated').text(item.dateTimeCreated && item.dateTimeCreated.toLocaleString());
    $('#dateTimeModified').text(item.dateTimeModified && item.dateTimeModified.toLocaleString());
    $('#itemClass').text(item.itemClass);
    $('#itemId').text(item.itemId);
    $('#itemType').text(item.itemType);

    item.body.getAsync('html', function(result){
      if (result.status === 'succeeded') {
        $('#bodyHtml').text(result.value);
      }
    });

    item.body.getAsync('text', function(result){
      if (result.status === 'succeeded') {
        $('#bodyText').text(result.value);
      }
    });

    if (item.itemType == Office.MailboxEnums.ItemType.Message){
      loadMessageProps(item);
    }
  }

  function errorHandler(error) {
         showNotification(error);
     }

 // Display notifications in message banner at the top of the task pane.
 function showNotification(content) {
   app.showNotification('Debug', content);
 }
