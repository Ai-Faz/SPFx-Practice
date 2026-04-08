import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';

import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './DemoHelloWorldWebPart.module.scss';
import * as strings from 'DemoHelloWorldWebPartStrings';
import MockHttpClient from './MockHttpClient';

// Props
export interface IDemoHelloWorldWebPartProps {
  description: string;
  test: string;
  test1: boolean;
  test2: string;
  test3: boolean;
}

// Models
export interface ISPList {
  Title: string;
  Id: string;
}

export interface ISPLists {
  value: ISPList[];
}

export default class DemoHelloWorldWebPart extends BaseClientSideWebPart<IDemoHelloWorldWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    this.domElement.innerHTML = `
      <section class="${styles.demoHelloWorld}">
        <div class="${styles.welcome}">
          <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>

          <p>Site: ${escape(this.context.pageContext.web.title)}</p>
          <p>User: ${escape(this.context.pageContext.user.displayName)}</p>

          <div id="spListContainer"></div>
        </div>
      </section>
    `;
     this._renderListAsync();
  }

  // Get Mock Data
  private _getMockListData(): Promise<ISPLists> {
    return MockHttpClient.get()
      .then((data: ISPList[]) => {
        return { value: data };
      });
  }

  // Render UI
  private _renderList(items: ISPList[]): void {
    let html: string = '';

    items.forEach((item: ISPList) => {
      html += `
        <ul class="${styles.list}">
          <li class="${styles.listItem}">
            <span>${item.Title}</span>
          </li>
        </ul>`;
    });

    const listContainer = this.domElement.querySelector('#spListContainer');
    if (listContainer) {
      listContainer.innerHTML = html;
    }
  }

  // Async Call
    private _renderListAsync(): void {
      this._getMockListData()
        .then((response) => {
          this._renderList(response.value);
        })
        .catch((error) => {
          console.error("Error fetching list:", error);
        });
    }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    return Promise.resolve(
      this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentSharePoint
        : strings.AppSharePointEnvironment
    );
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;
    this._isDarkTheme = !!currentTheme.isInverted;
  }

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