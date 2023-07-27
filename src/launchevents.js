Office.onReady()

function eventHandlers(event) {
  setInterval(() => {
    console.log('working')
  }, 100)
}

function onMessageSendHandler(event) {
  event.completed({ allowEvent: false, errorMessage: '' })
}

Office.actions.associate('onMessageComposeHandler', eventHandlers)
Office.actions.associate('onMessageSendHandler', onMessageSendHandler)
