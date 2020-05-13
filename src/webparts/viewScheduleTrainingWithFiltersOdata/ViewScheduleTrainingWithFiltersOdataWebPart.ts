import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ViewScheduleTrainingWithFiltersOdataWebPart.module.scss';
import * as strings from 'ViewScheduleTrainingWithFiltersOdataWebPartStrings';
import * as pnp from 'sp-pnp-js';
import { Web } from 'sp-pnp-js';

export interface IViewScheduleTrainingWithFiltersOdataWebPartProps {
  description: string;
  site:string;
  listTitle:string;
}

export default class ViewScheduleTrainingWithFiltersOdataWebPart extends BaseClientSideWebPart <IViewScheduleTrainingWithFiltersOdataWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <table>
      <tr>
        <td>Title</td>
        <td>:</td>
        <td><input type='text' class='title' /></td>
      </tr>
      <tr>
        <td>Location</td>
        <td>:</td>
        <td><input type='text' class='location' /></td>
      </tr>
      <tr>
        <td colspan='3'><input type='button' class='btnSearch' value='Searhc' /></td>
      </tr>
    </table>
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

    this.domElement.querySelector('.btnSearch').addEventListener('click', () => { this.searchSchedule(); });
  }

  private searchSchedule(): void{
    var title = (this.domElement.querySelector('.title') as HTMLInputElement).value;
    var location = (this.domElement.querySelector('.location') as HTMLInputElement).value;

    var htmlString = "";

    this.GetTrainingResult(title, location).then(items => {
      items.forEach(item => {
          htmlString += "<tr>";
          htmlString += "<td>" + item.Title + "</td>";
          htmlString += "<td>" + item.Location + "</td>";
          htmlString += "<td>" + item.EventDate + "</td>";
          htmlString += "<td>" + item.EndDate + "</td>";
          htmlString += "</tr>";
      })

      this.domElement.querySelector('listdata-training-events').innerHTML = htmlString;
    })
  }

  private GetTrainingResult(title : string, location : string): Promise<any[]>{
    const _web = new pnp.Web(this.properties.site);

    if((typeof this.properties.site != 'undefined' && this.properties.site != '') && (typeof this.properties.listTitle != 'undefined' && this.properties.listTitle != ''))
    {
      if(title != '' && location != '')
        return _web.lists.getByTitle(this.properties.listTitle).items.filter("substringof('" + title + "', Title) and substringof('" + location + "', Location)").get();
      else if(title != '' && location == '')
        return _web.lists.getByTitle(this.properties.listTitle).items.filter("substringof('" + title + "', Title)").get();
      else if(title == '' && location != '')
        return _web.lists.getByTitle(this.properties.listTitle).items.filter("substringof('" + location + "', Location)").get();
      else
        return _web.lists.getByTitle(this.properties.listTitle).items.get();

    }
    else
    {
      alert('Please set properties first');
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
