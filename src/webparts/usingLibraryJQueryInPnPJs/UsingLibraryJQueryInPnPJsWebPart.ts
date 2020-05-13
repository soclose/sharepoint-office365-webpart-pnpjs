import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './UsingLibraryJQueryInPnPJsWebPart.module.scss';
import * as strings from 'UsingLibraryJQueryInPnPJsWebPartStrings';

import { SPComponentLoader } from '@microsoft/sp-loader';
import * as $ from 'jQuery';

import { Web, Item } from 'sp-pnp-js';
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions
} from '@microsoft/sp-http';
import ViewScheduleTrainingWithFiltersOdataWebPart from '../viewScheduleTrainingWithFiltersOdata/ViewScheduleTrainingWithFiltersOdataWebPart';

export interface IUsingLibraryJQueryInPnPJsWebPartProps {
  description: string;
  SiteUrl: string;
  ListLibrary: string;
}

export default class UsingLibraryJQueryInPnPJsWebPart extends BaseClientSideWebPart <IUsingLibraryJQueryInPnPJsWebPartProps> {

  jQuery: any;

  public render(): void {
    this.domElement.innerHTML = `
    <div class="carousel slide" id="listJQueryCarousel" data-ride="carousel">
    </div>
    `;

     SPComponentLoader.loadScript('https://ajax.googleapis.com/ajax/libs/jquery/3.3.1/jquery.min.js').then(($: any): void => {
      this.jQuery = $;
      SPComponentLoader.loadScript('https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js').then((): void => {
        this.JQueryCarousel();
      });
    });
  }

  
  private JQueryCarousel(): void{
    const _web = new Web(this.properties.SiteUrl);
    
    var carosoulIndicators = '';
    var displayItems = '';
    var oCarousel = [];

    carosoulIndicators += '<ol class="carousel-indicators">';
    displayItems += '<div class="carousel-inner">';
    var i = 0;

    try
    {
      let requests: Array<Promise<SPHttpClientResponse>> = new Array <Promise<SPHttpClientResponse>>();

      if(this.properties.ListLibrary != '') {
        const _request = _web.lists.getByTitle(this.properties.ListLibrary).items.select("Title", "ID", "Created", "AttachmentFiles", "EncodedAbsUrl")
        .expand("AttachmentFiles")
        .get();
  
        requests.push(_request);
        let oRequests = new Array();

        Promise.all(requests).then((responses: SPHttpClientResponse[]) => {
          responses.forEach(response => { 
            oRequests.push(response);
          });

          Promise.all(oRequests).then((oResponse: any[]) => {
            for(var index = 0; index < oResponse.length; index++){
              for(var j = 0; j < oResponse[i].length; j++){
                oCarousel.push(oResponse[i][j]);
              }
            }
            oCarousel.forEach(item => {
              if(i == 0){
                carosoulIndicators += '<li data-target="#carousel-example-generic" data-slide-to="' + i + '" class="active"></li>';  
                displayItems += '<div class="item active">';
                if(typeof item.AttachmentFiles[0] != "undefined"){
                  displayItems += '<img src="' + item.AttachmentFiles[0].ServerRelativeUrl + '" alt="' + item.Title + '" style="display: block;margin-left: auto;margin-right: auto;width: 75%;height:80%">';
                }
                displayItems += '<a target="_blank" href="' + item.EncodedAbsUrl.replace("1_.000", "") + '/DispForm.aspx?ID=' + item.ID + '"><div class="carousel-caption">';
                displayItems += item.Title;
                displayItems += '</div></a></div>';
                
              } else {
                carosoulIndicators += '<li data-target="#carousel-example-generic" data-slide-to="' + i + '" class=""></li>';
                displayItems += '<div class="item">';
                if(typeof item.AttachmentFiles[0] != "undefined"){
                  displayItems += '<img src="' + item.AttachmentFiles[0].ServerRelativeUrl + '" alt="' + item.Title + '" style="display: block;margin-left: auto;margin-right: auto;width: 75%;height:80%">';
                }
                displayItems += '<a target="_blank" href="' + item.EncodedAbsUrl.replace("1_.000", "") + '/DispForm.aspx?ID=' + item.ID + '"><div class="carousel-caption">';
                displayItems += item.Title;
                displayItems += '</div></a></div>';
    
              }
              i++;
            });
  
            displayItems += '</div>';
            carosoulIndicators += '</ol>';
    
            displayItems = carosoulIndicators + displayItems;
            displayItems += '<a class="left carousel-control" href="#carousel-example-generic" data-slide="prev">';
            displayItems += '<span class="fa fa-angle-left"></span>';
            displayItems += '</a>';
            displayItems += '<a class="right carousel-control" href="#carousel-example-generic" data-slide="next">';
            displayItems += '<span class="fa fa-angle-right"></span>';
            displayItems += '</a>';
            
            $("#listJQueryCarousel", document.body).html(displayItems); 
            (<any>jQuery(".carousel", document.body)).carousel();

          })
        })
      }
    }
    catch(e)
    {
      console.error(e);
    }
  }
  

  protected get dataVersion(): Version {
  return Version.parse('1.0');
}

protected onInit(): Promise<void> {
  SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/font-awesome/4.7.0/css/font-awesome.min.css');
  SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css');

  return Promise.resolve(undefined);
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
              }),
              PropertyPaneTextField('SiteUrl', {
                label: "Site Url"
              }),
              PropertyPaneTextField('ListLibrary', {
                label: "List Library"
              })
            ]
          }
        ]
      }
    ]
  };
}
}
