import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';

import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './DemoHelloWorldWebPart.module.scss';
import * as strings from 'DemoHelloWorldWebPartStrings';  

//  Mock Client
import MockHttpClient from './MockHttpClient';

//  SharePoint HTTP Client
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

// 
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';

// ===================== INTERFACES =====================

export interface IDemoHelloWorldWebPartProps {
  description: string;
  test: string;
  test1: boolean;
  test2: string;
  test3: boolean;
}

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
}

// ===================== MAIN CLASS =====================

export default class DemoHelloWorldWebPart
  extends BaseClientSideWebPart<IDemoHelloWorldWebPartProps> {

  // ===================== RENDER =====================
  public render(): void {

    this.domElement.innerHTML = `
      <section class="${styles.demoHelloWorld}">
        <div class="${styles.welcome}">
          <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
          <p>Site: ${escape(this.context.pageContext.web.title)}</p>
          <p>User: ${escape(this.context.pageContext.user.displayName)}</p>
        </div>
        <div id="spListContainer"></div>
      </section>
    `;

    console.log("Render called");
    this._renderListAsync();
  }

  // ===================== MOCK DATA =====================
  private _getMockListData(): Promise<ISPLists> {
    return MockHttpClient.get()
      .then((data: ISPList[]) => {
        return { value: data };
      });
  }

  // ===================== SHAREPOINT DATA =====================
  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient
      .get(
        this.context.pageContext.web.absoluteUrl +
        `/_api/web/lists?$filter=Hidden eq false`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  // ===================== RENDER LIST =====================
  private _renderList(items: ISPList[]): void {

    let html: string = '';

    items.forEach((item: ISPList) => {
      html += `
        <ul class="${styles.list}">
          <li class="${styles.listItem}">
            <span>${item.Title}</span>
          </li>
        </ul>
      `;
    });

    const container = this.domElement.querySelector('#spListContainer');

    //  Null safety
    if (container) {
      container.innerHTML = html;
    }
  }

  // ===================== MAIN LOGIC =====================
  private _renderListAsync(): void {

    //  Local Workbench
    if (Environment.type === EnvironmentType.Local) {

      this._getMockListData()
        .then((response) => {
          console.log("Mock Data:", response);
          this._renderList(response.value);
        })
        .catch((error) => console.error("Mock Error:", error));

    }

    //  SharePoint Workbench
    else if (
      Environment.type === EnvironmentType.SharePoint ||
      Environment.type === EnvironmentType.ClassicSharePoint
    ) {

      this._getListData()
        .then((response) => {
          console.log("SP Data:", response);
          this._renderList(response.value);
        })
        .catch((error) => console.error("SP Error:", error));
    }
  }

  // ===================== OTHER =====================
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', { label: strings.DescriptionFieldLabel }),
                PropertyPaneTextField('test', { label: 'Text', multiline: true }),
                PropertyPaneCheckbox('test1', { text: 'Checkbox' }),
                PropertyPaneDropdown('test2', {
                  label: 'Dropdown',
                  options: [
                    { key: '1', text: 'One' },
                    { key: '2', text: 'Two' },
                    { key: '3', text: 'Three' }
                  ]
                }),
                PropertyPaneToggle('test3', {
                  label: 'Toggle',
                  onText: 'On',
                  offText: 'Off'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}