# PowerPoint to Places

PowerPoint to Places is a tool that will automate the sending of text to destinations during a PowerPoint talk. The original objective was that it would post messages via lack WebHooks](https://api.slack.com/messaging/webhooks) but it's got enough flexibility that other places (such as Teams, Discord or Twitter) could be used.

## How it works

Edit the `SlackPlace.cs` file and add the URL for your webhook.

Then in PowerPoint, add lines to the **notes** for a slide that starts with `SLACK:` and when the app runs (and you're presenting in PowerPoint), that line will be sent to Slack via the webhook.

## Can it send to `<somewhere>`?

Sure, implement the `ISendToPlace` interface and register it in the DI container, then you're good to go. Maybe even send me a PR so others can use it!

## Caveats

You need to be on Windows.

No, there isn't a pre-compiled "release", you have to compile it with VS 2019.

This is doing COM interop with PowerPoint, so you need PowerPoint desktop.

I made it for my scenario, so I can't be sure it'll work for you.

Yes, you can spam a Slack channel a lot.
