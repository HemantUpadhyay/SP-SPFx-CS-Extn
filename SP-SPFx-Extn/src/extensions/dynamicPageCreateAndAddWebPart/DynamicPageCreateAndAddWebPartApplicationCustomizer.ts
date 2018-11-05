import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'DynamicPageCreateAndAddWebPartApplicationCustomizerStrings';

const LOG_SOURCE: string = 'DynamicPageCreateAndAddWebPartApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IDynamicPageCreateAndAddWebPartApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class DynamicPageCreateAndAddWebPartApplicationCustomizer
  extends BaseApplicationCustomizer<IDynamicPageCreateAndAddWebPartApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    // Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    // let message: string = this.properties.testMessage;
    // if (!message) {
    //   message = '(No properties were provided.)';
    // }

    // Dialog.alert(`Hello from ${strings.Title}:\n\n${message}`);
    this.Initiate();
    return Promise.resolve();
  }

  @override
  public onRender(): void{
      
  }

  private async getNewPageStatus() {
    // this.context.aadHttpClientFactory
    //   .getClient('https://tenant.onmicrosoft.com/6b347c27-f360-47ac-b4d4-af78d0da4223')
    //   .then((client: AadHttpClient): void => {
    //     client
    //       .get('https://myfunction.azurewebsites.net/api/CurrentUser', AadHttpClient.configurations.v1)
    //       .then((response: HttpClientResponse): Promise<JSON> => {
    //         return response.json();
    //       })          
    //   });

    const currentWebUrl: string = this.context.pageContext.web.absoluteUrl;
    //const pageName = 'DynamicPage.aspx'
    var functionUrl : string = "https://[functionName].azurewebsites.net/api/[FunctionMethod1]";

    var functionPageCreation : string = "https://[functionName].azurewebservices.net/api/[FunctionMethod2]";
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    requestHeaders.append("Cache-Control", "private"); 

    const postOptions : RequestInit = {
    headers: requestHeaders,
    //body: `{\r\n    siteURL: '${currentWebUrl}',\r\n    pageName: '${pageName}' \r\n}`,
    body: `{\r\n    siteURL: '${currentWebUrl}'\r\n}`,
    method: "POST"
    };
    let responseText: string = "";
    let createPageStatus: string = "";
    console.log("Wait started for Creating page");
    await fetch(functionUrl, postOptions).then((response) => {
        console.log("Response returned");
        if (response.ok) {
          return response.json()


        }
        else
        {
            var errMsg = "Error detected while adding site page. Server response wasn't OK ";
            console.log(errMsg);
        } 
      }).then((responseJSON: JSON) => {
        responseText = JSON.stringify(responseJSON).trim();
        console.log(responseText);
        if(responseText.toLowerCase().indexOf("success") > 0)
            {
              console.log("success feedback");
            }
        if(responseText.toLowerCase().indexOf("error") > 0)
            {
              console.log("web call errored");
            }
      }
    ).catch ((response: any) => {
      let errMsg: string = `WARNING - error when calling URL ${functionUrl}. Error = ${response.message}`;
      console.log(errMsg);
    });
    console.log("wait finished");
  }

  private async Initiate() {
    await this.getNewPageStatus();
  }
}
