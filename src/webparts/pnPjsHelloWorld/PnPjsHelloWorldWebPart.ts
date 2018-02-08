import { Version,
      Environment,
      EnvironmentType
 } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

//import { EnvironmentType } from '@microsoft/sp-client-base';

import { escape } from '@microsoft/sp-lodash-subset';

import styles from './PnPjsHelloWorldWebPart.module.scss';
import * as strings from 'PnPjsHelloWorldWebPartStrings';

import MockHttpClient from './MockHttpClient'
import pnp from 'sp-pnp-js';

export interface IPnPjsHelloWorldWebPartProps {
  description: string;
}

export default class PnPjsHelloWorldWebPart extends BaseClientSideWebPart<IPnPjsHelloWorldWebPartProps> {

  private _renderListAsync(): void {
    // Local environment
    if (Environment.type === EnvironmentType.Local) {
      this._getMockListData().then((response) => {
        this._renderList(response);
      }); }
      //SharePoint Site environment
      else {
      this._getListData()
        //.then((response) => {
        // this._renderDocuments(response);
        //});
    }
  }

  private _renderList(items: ISPList[]): void {
    let html: string = '';
    items.forEach((item: ISPList) => {
      html += `
      <ul class="${styles.list}">
          <li class="${styles.listItem}">
              <span class="ms-font-l">${item.Title}</span>
          </li>
      </ul>`;
    });
    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;
  }

  private _renderDocuments(items: any[]): void {
      let html: string = '';
      items.forEach((item: any) => {
        if(item.ContentTypeId != "0x01200075D313699C70804495D0C04BF909CA71")
        {
          html += `
            <ul class="${styles.list}">
              <li class="${styles.listItem}">
                <span class="ms-font-1">Site Name - ${item.SiteName}</span><br>
                <span class="ms-font-1">Project ID - ${item.ProjectID}</span> <br>
                <span class="ms-font-1">Sheet Number - ${item.SheetNumber}</span><br>
                <span class="ms-font-1">Discipline - ${item.Discipline}</span><br>
                <span class="ms-font-1"><a href="${item.odata.editLink}">Click Here to View the Document</a></span>
              </li>
            </ul>
          `
        }
      });
      const listContainer: Element = this.domElement.querySelector('#spListContainer');
      listContainer.innerHTML = html;
    }

  private _getListData(): void {
    pnp.sp.web.lists.getByTitle('TestFileRepository').items.get().then((items: any[]) => {
      console.log(items);
      this._renderDocuments(items);
    });
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.pnPjsHelloWorld }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
            </div>
          </div>
          <div id='spListContainer' />
        </div>
      </div>`;

      this._renderListAsync();
  }

  private _getMockListData(): Promise<ISPList[]> {
    return MockHttpClient.get(this.context.pageContext.web.absoluteUrl).then(() => {
      const listData: ISPList[] = [
        { Title: 'Mock List', Id: '1' },
        { Title: 'Mock List Two', Id: '2' },
        { Title: 'Mock List Three', Id: '3' }
      ];
      return listData;
    }) as Promise<ISPList[]>;
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

export interface ISPList {
  Title: string;
  Id: string;
}
