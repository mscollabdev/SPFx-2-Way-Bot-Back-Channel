# SharePoint Framework Bot with Bi-directional Context via Back Channel
Simple SharePoint Framework Project that embeds a bot and uses the Bot Framework back channel to silently integrate the bot and client to share contextual information and API calls.

## Setup
The SharePoint Framework leverages a number of tools including Gulp, so I recommend setting up your environment with the following the [Set up your SharePoint client-side web part development environment](https://dev.office.com/sharepoint/docs/spfx/set-up-your-development-environment) on dev.office.com.

## Running the sample
Run the following commands from a command prompt

<pre>
npm install
gulp serve
</pre>

When the gulp serve process completes, the local SharePoint Workbench will be launched. Because this bot performs queries against SharePoint REST APIs, you need to change the URL to the hosted SharePoint Workbench (https://tenant.sharepoint.com/_layouts/15/workbench.aspx).

## Back channel logic
All of the application logic related to using the back channel is located in the [EchoBotWebPart.ts](https://github.com/richdizz/SPFx-2-Way-Bot-Back-Channel/blob/master/src/webparts/echoBot/EchoBotWebPart.ts) file.