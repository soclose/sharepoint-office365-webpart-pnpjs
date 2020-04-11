import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ViewScheduleTrainingWebPart.module.scss';
import * as strings from 'ViewScheduleTrainingWebPartStrings';
import * as pnp from 'sp-pnp-js';
import { Web } from 'sp-pnp-js';

export interface IViewScheduleTrainingWebPartProps {
  description: string;
  site:string;
  listTitle:string;
}

export default class ViewScheduleTrainingWebPart extends BaseClientSideWebPart <IViewScheduleTrainingWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <table>
        <thead>
          <tr>
            <td>Title</td>
            <td>Location</td>
            <td>Event Date</td>
            <td>End Date</td>
          </tr>
        </thead>
        <tbody class='listdata-training-events'>
        </tbody>
      </table>
    `;

    this.GetTrainingEvents();
  }

  private GetTrainingEvents(): void{
    if((typeof this.properties.site != 'undefined' && this.properties.site != '') && (typeof this.properties.listTitle != 'undefined' && this.properties.listTitle != ''))
    {
      var htmlString = "";
      const _web = new pnp.Web(this.properties.site);
      _web.lists.getByTitle(this.properties.listTitle).items.get().then(response => {
        response.forEach(item => {
          debugger
          htmlString += "<tr>";
          htmlString += "<td>" + item.Title + "</td>";
          htmlString += "<td>" + item.Location + "</td>";
          htmlString += "<td>" + item.EventDate + "</td>";
          htmlString += "<td>" + item.EndDate + "</td>";
          htmlString += "</tr>";
        })

        this.domElement.querySelector(".listdata-training-events").innerHTML = htmlString;
      })
    }
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
              }),
              PropertyPaneTextField('site', {
                label:'Site Name'
              }),
              PropertyPaneTextField('listTitle', {
                label:'List Library'
              })
            ]
          }
        ]
      }
    ]
  };
}
}
