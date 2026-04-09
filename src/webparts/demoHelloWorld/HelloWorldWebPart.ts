// import { Version } from '@microsoft/sp-core-library';

// import {
// IPropertyPaneConfiguration,
// PropertyPaneTextFie1d,
// // add a few more properties to the property pane
// PropertyPaneCheckbox,
// PropertyPaneDropdown,
// PropertyPaneTogg1e
// } from '@microsoft/sp-property-pane' ;

// import { BaseC1ientSideWebPart } from '@microsoft/sp-webpart-base';
// import { escape, fromPairs } from '@microsoft/sp-lodash-subset';
// import styles from './DemoHelloWorldWebPart.module.scss';
// import * as strings from 'HelloWorldWebPartStrings';

// // Import mock Http Client Module 

// import MockHttpClient from './MockHttpClient';

// //Helper Class to execute REST API Request agesnst SharePoint 
// // /_api/web/lists

// import {
//     SPHttpClient,
//     SPHttpclientResponse
// }from '@microsoft/sp-http';

// // Check Enviroment Type
// import {
//     Enviroment,
//     EnviromentType
// }from '@microsoft/sp-core-library';


// // Props
// export interface IDemoHelloWorldWebPartProps {
//   description: string;
//   test: string;
//   test1: boolean;
//   test2: string;
//   test3: boolean;
// }

// // Defile List Models

// export interface ISPLists {
//   value: ISPList[];
// }
// export interface ISPList {
//   Title: string;
//   Id: string;
// }
// export default class HelloWorldWebPart extends BaseC1ientSideWebPart<IDemoHelloWorldWebPartProps> {
// public render() : void {
//     this.domElement.innerHtml = '
// public
// this. domE1ement.innerHTML =
// <div class="${ styles. helloWor1d }">
// <div styles. container
// <div class="${ styles. row }">
// <div class="${ styles. column
// <span class="${ styles. title }">SharePoint With Darwish</span>
// class="${ styles.subTit1e }">SharePoint SPFx.</p>
// class="${ styles.description
// } " >${escape(this . properties . description)
// class="${ styles. description
// } " >${escape(this . properties. test)
//     '
// }
// }