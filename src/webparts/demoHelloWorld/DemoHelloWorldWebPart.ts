import { Version } from '@microsoft/sp-core-library';

import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  // Adding few more Propertyes Here 
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';

import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './DemoHelloWorldWebPart.module.scss';
import * as strings from 'DemoHelloWorldWebPartStrings';

export interface IDemoHelloWorldWebPartProps {
  description: string;
  //Adding more Propertys 
  test : string;
  test1 : boolean;
  test2 : string;
  test3 : boolean;
}

export default class DemoHelloWorldWebPart extends BaseClientSideWebPart<IDemoHelloWorldWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = ''; 

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.demoHelloWorld} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>

          <p class="${styles.description}">Title ${escape(this.context.pageContext.web.title)}</p>
          <p class="${styles.description}">Site Language ${escape(this.context.pageContext.web.languageName)}</p>
          <p class="${ styles.description }">User ${escape(this.context.pageContext.user.displayName)}</p>
          <p class="${ styles.description }">LoginName ${escape(this.context.pageContext.user.loginName)}</p>
          <p class="${ styles.description }">URL ${escape(this.context.pageContext.site.absoluteUrl)}</p>

          <p class="${styles.description}">${escape(this.properties.test)}</p>

          <p class="${styles.description}">Dropdown Value is : ${escape(this.properties.test2)}</p>

          <a href="https://aka.ms/spfx" class="${styles.button}">
          <span>Learn More</span></a>
      </div>
    </section>`;
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
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
                PropertyPaneTextField('test',{
                  label:'Multi-Line Text Field',
                  multiline:true
                }),
                PropertyPaneCheckbox('test1',{
                  text: 'Checkbox'
                }),
                PropertyPaneDropdown('test2',{
                  label:'Dropdown',
                  options : [
                    {key:'1' , text:'one'},
                    {key:'2' , text:'two'},
                    {key:'3' , text:'three'}
                  ]}),
                  PropertyPaneToggle('test3',{
                    label: 'toggel',
                    onText: 'on',
                    offText: 'off'
                  })
              ]
            }
          ]
        }
      ]
    };
  }
}
