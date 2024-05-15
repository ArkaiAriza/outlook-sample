Office.onReady()

function eventHandlers(event) {
  console.log('preprocess')
  event.completed()
}

function onMessageSendHandler(event) {
  console.log('sending')
  event.completed({ allowEvent: true })
}

Office.actions.associate('onMessageComposeHandler', eventHandlers)
Office.actions.associate('onMessageSendHandler', onMessageSendHandler)
