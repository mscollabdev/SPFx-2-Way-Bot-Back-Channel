import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import { App } from '../../BotFramework-WebChat/botchat';
import { DirectLine } from 'botframework-directlinejs';
require('../../../src/BotFramework-WebChat/botchat.css');
import * as strings from 'echoBotStrings';
import { IEchoBotWebPartProps } from './IEchoBotWebPartProps';

export default class EchoBotWebPart extends BaseClientSideWebPart<IEchoBotWebPartProps> {

  public render(): void {
    // Generate a random element id for the WebChat container
    var possible:string = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
    var elementId:string = "";
    for(var i = 0; i < 5; i++)
      elementId += possible.charAt(Math.floor(Math.random() * possible.length));
    this.domElement.innerHTML = '<div id="' + elementId + '"></div>';

    // Get userprofile from SharePoint REST endpoint
    var req = new XMLHttpRequest();
    req.open("GET", "/_api/SP.UserProfiles.PeopleManager/GetMyProperties", false);
    req.setRequestHeader("Accept", "application/json");
    req.send();
    var user = { id: "userid", name: "unknown" };
    if (req.status == 200) {
      var result = JSON.parse(req.responseText);
      user.id = result.Email;
      user.name = result.DisplayName;
    }

    // Initialize DirectLine connection
    var botConnection = new DirectLine({
      secret: "5UCju9ytkk4.cwA.Uk4.1GB38_MVXofdjsxbPeVMixHda-3DUMNmQSJzq5_ahrI"
    });

    // Initialize the BotChat.App with basic config data and the wrapper element
    App({
        user: user,
        botConnection: botConnection
      }, document.getElementById(elementId));

    // Call the bot backchannel to give it user information
    botConnection
      .postActivity({ type: "event", name: "initialize", value: user.name, from: user })
      .subscribe(id => console.log("success initializing"));

    // Listen for events on the backchannel
    botConnection.activity$
      .subscribe(a => {
        var activity:any = a;
        if (activity.type == "event" && activity.name == "runShptQuery") {
          // Parse the entityType out of the value query string
          var entityType = activity.value.substr(activity.value.lastIndexOf("/") + 1);

          // Perform the REST call against SharePoint
          var shptReq = new XMLHttpRequest();
          shptReq.open("GET", activity.value, false);
          shptReq.setRequestHeader("Accept", "application/json");
          shptReq.send();
          var shptResult = JSON.parse(shptReq.responseText);

          // Call the bot backchannel to give the aggregated results
          botConnection
            .postActivity({ type: "event", name: "queryResults", value: { entityType: entityType, count: shptResult.value.length }, from: user })
            .subscribe(id => console.log("success sending results"));
        }
      });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
