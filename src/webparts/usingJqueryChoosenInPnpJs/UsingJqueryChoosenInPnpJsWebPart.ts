import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './UsingJqueryChoosenInPnpJsWebPart.module.scss';
import * as strings from 'UsingJqueryChoosenInPnpJsWebPartStrings';

import * as $ from 'jQuery';
import { SPComponentLoader } from '@microsoft/sp-loader';

export interface IUsingJqueryChoosenInPnpJsWebPartProps {
  description: string;
}

export default class UsingJqueryChoosenInPnpJsWebPart extends BaseClientSideWebPart <IUsingJqueryChoosenInPnpJsWebPartProps> {

  jQuery : any;

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.usingJqueryChoosenInPnpJs }">
      <div class="ms-Dropdown" tabindex="0">
        <label class="ms-Label">Country</label>
        <select data-placeholder="Choose a country..." class="sample-dropdown">
          <option>--- Please Select ---</option>
          <option>Indonesia</option>
          <option>USA</option>
          <option>England</option>
          <option>France</option>
        </select>
        <br>
        <label class="ms-Label">City&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</label>
        <select data-placeholder="Choose a city..." class="sample-dropdown-city">
          <option>--- Please Select ---</option>
        </select>
      </div>
      </div>`;

      SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/jquery/3.3.1/jquery.min.js').then(($: any): void => {
        this.jQuery = $;
        SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/chosen/1.8.7/chosen.jquery.js').then((): void => {
          (<any>jQuery(".sample-dropdown", document.body)).chosen({ search_contains : true });
          (<any>jQuery(".sample-dropdown-city", document.body)).chosen({ search_contains : true });

          const dropdownCountry = document.querySelector('.sample-dropdown').nextSibling;
          
          dropdownCountry.addEventListener('click', (event) => {  
            this.restructureCity(event);
          });
        });
      });

  }

  private restructureCity(event): void{
    var params = event.target.textContent;
    
    var optionsCity = "";
    if(params == "Indonesia"){
      optionsCity = "";
      optionsCity += "<option>--- Please Select ---</option>";
      optionsCity += "<option>Jakarta</option>";
      optionsCity += "<option>Bandung</option>";
      optionsCity += "<option>Surabaya</option>";
      $(".sample-dropdown-city", document.body).html(optionsCity);
    } else if(params == "USA"){
      optionsCity = "";
      optionsCity += "<option>--- Please Select ---</option>";
      optionsCity += "<option>Washington, D.C.</option>";
      optionsCity += "<option>New York</option>";
      optionsCity += "<option>Sun Francisco</option>";
      $(".sample-dropdown-city", document.body).html(optionsCity);
    } else if(params == "England"){
      optionsCity = "";
      optionsCity += "<option>--- Please Select ---</option>";
      optionsCity += "<option>Manchester</option>";
      optionsCity += "<option>London</option>";
      optionsCity += "<option>Liverpool</option>";
      $(".sample-dropdown-city", document.body).html(optionsCity);
    } else if(params == "France"){
      optionsCity = "";
      optionsCity += "<option>--- Please Select ---</option>";
      optionsCity += "<option>Paris</option>";
      optionsCity += "<option>Marseille</option>";
      optionsCity += "<option>Lyon</option>";
      $(".sample-dropdown-city", document.body).html(optionsCity);  
    } else {
      optionsCity = "";
      optionsCity += "<option>--- Please Select ---</option>";
      $(".sample-dropdown-city", document.body).html(optionsCity); 
    }
    
    (<any>jQuery(".sample-dropdown-city", document.body)).trigger("chosen:updated");
    
  }

  protected onInit(): Promise<void> {
    SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/chosen/1.8.7/chosen.min.css');
  
    return Promise.resolve(undefined);
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
