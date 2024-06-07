Office.onReady();

function eventHandler(event) {
  console.log('preprocess');
  event.completed();
}

function recipientsHandler(event) {
  console.log('recipients');
  event.completed();
}

function onMessageSendHandler(event) {
  console.log('sending');
  event.completed({ allowEvent: true });
}

Office.actions.associate('onMessageComposeHandler', eventHandler);
Office.actions.associate(
  'onMessageRecipientsChangedHandler',
  recipientsHandler
);

Office.actions.associate('onMessageSendHandler', onMessageSendHandler);
