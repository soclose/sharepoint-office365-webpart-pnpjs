import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneChoiceGroup
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';

export interface IHelloWorldWebPartProps {
  description: string;
  myClass: string;
  enableStudy:string;
  score:string;
  choice:string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart <IHelloWorldWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      Hello World My Name is Nurjalih <br>
      i'm Class ${escape(this.properties.myClass)} <br>
      Enable Study ${escape(this.properties.enableStudy)} <br>
      My Score ${escape(this.properties.score)} <br>
      I'm ${escape(this.properties.choice)}
    `;
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
            groupName: 'Custom Properties Nurjalih',
            groupFields: [
              PropertyPaneTextField('description', {
                label: 'Description Web Part'
              }),
              PropertyPaneDropdown('myClass', {
                label: 'Class',
                options: [
                  {key : '1', text: 'Class 1'},
                  {key : '2', text: 'Class 2'},
                  {key : '3', text: 'Class 3'}
                ]
              }),
              PropertyPaneToggle('enableStudy', {
                label: 'Enable my Study ?'
              }),
              PropertyPaneSlider('score', {
                label:'SCore', min:1, max:100
              }),
              PropertyPaneChoiceGroup('choice', {
                label:'Choice',
                options:[
                  { key: 'Male', text:'Male'},
                  { key: 'Female', text:'Female'}
                ]
              })
            ]
          }
        ]
      }
    ]
  };
}
}
