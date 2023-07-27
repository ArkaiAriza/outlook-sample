Small code sample to reproduce an issue in Outlook Web Add-Ins

## Installation

```bash
yarn install
```

```bash
yarn start
```

## Usage

Once the server is running sideload the add-in by going to https://localhost:3000/manifest.xml and save the file. Then go to outlook web and click in get add ins, and add a custom add-in from the file you just saved.

This add-in contains handlers for two events: OnMessageCompose and OnMessageSend. The OnMessageCompose handler will start an interval and log the word 'working' every 100ms. The OnMessageSend handler will complete the event and block the message being sent with smart alerts. This is to demonstrate that when completing the event in the OnMessageSend handler, the OnMessageCompose handler also stops it's execution

Note: You might get a rollup error about requiring an input. Please ignore it, as it is not necesary for this example to work.