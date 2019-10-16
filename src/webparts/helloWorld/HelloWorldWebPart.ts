import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';

import MockHttpclient from './MockHttpClient';

import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';

export interface IHelloWorldWebPartProps {
  description: string;
}

export interface ISPLists {
  value:ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <div class="${ styles.main }">
        <div class="${ styles.row }">
            <div class="${ styles.column} }">
                <div class="${ styles.headercontent }" onclick="window.open('http://google.com', '_blank');">
                    <p>Business process</p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="${ styles.headercontent}">
                    <p class="${ styles.doubleline}">Architecture & Applications</p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="${ styles.headercontent}">
                    <p>Project</p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="${ styles.headercontent}">
                    <p>Release</p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="${ styles.headercontent}">
                    <p>Change Governance</p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="${ styles.headercontent}">
                    <p>Monitor</p>
                </div>
            </div>
        </div>


        <div class="${ styles.row}">
            <div class="${ styles.column}">
                <div class="${ styles.content}">
                    <p class="${ styles.doubleline}">Record To Report (R2R)</p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="${ styles.content}">
                    <p class="${ styles.doubleline}">Architecture & System Landscape</p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="${ styles.content}">
                    <p class="${ styles.doubleline}">Project Portfolio Overview</p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="${ styles.content}">
                    <p class="${ styles.doubleline}">Release Management process</p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="${ styles.content}">
                    <p class="${ styles.doubleline}">Change Control Process - CAB</p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="${ styles.content}">
                    <p class="${ styles.doubleline}">SAP Technical Monitoring</p>
                </div>
            </div>
        </div>

        <div class="${ styles.row}">
            <div class="${ styles.column}">
                <div class="${ styles.content}">
                    <p>Procure To Pay (P2P)</p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="${ styles.content}">
                    <p>Application Catalog</p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="${ styles.content}">
                    <p class="${ styles.doubleline}">Project Management Framework</p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="${ styles.content}">
                    <p>Release Schedule</p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="${ styles.content}" style="padding-top:5px;padding-bottom:5px;">
                    <p class="${styles.tripleline}">Production Verification & Validation</p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="${ styles.content}">
                    <p class="${ styles.doubleline}">SAP Business Process Monitoring</p>
                </div>
            </div>
        </div>

        <div class="${ styles.row}">
            <div class="${ styles.column}">
                <div class="${ styles.content}">
                    <p class="${ styles.doubleline}">Schedule To Deliver (S2D)</p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="${ styles.content}">
                    <p>Interface Catalog</p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="${ styles.content}">
                    <p class="${ styles.doubleline}">Transition To Operations</p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="${ styles.content}">
                    <p>Testing Standards</p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="${ styles.content} ${ styles.empty}">
                    <p></p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="${ styles.content}">
                    <p>Control M</p>
                </div>
            </div>
        </div>

        <div class="${ styles.row}">
            <div class="${ styles.column}">
                <div class="${ styles.content}">
                    <p class="${ styles.doubleline}">Forecast To Produce (F2P)</p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="${ styles.content}">
                    <p>Integration Standards</p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="${ styles.content} ${ styles.empty}">
                    <p></p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="${ styles.content}">
                    <p>Test Catalog</p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="content ${ styles.empty}">
                    <p></p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="${ styles.content}">
                    <p>SloarWind</p>
                </div>
            </div>
        </div>

        <div class="${ styles.row}">
            <div class="${ styles.column}">
                <div class="${ styles.content}">
                    <p>Order To Cash (O2C)</p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="${ styles.content}">
                    <p>Coding Standards</p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="content ${ styles.empty}">
                    <p></p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="${ styles.content}">
                    <p class="${ styles.doubleline}">Solution Manager - CHaRM</p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="${ styles.content} ${ styles.empty}">
                    <p></p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="${ styles.content}">
                    <p>Network Monitoring</p>
                </div>
            </div>
        </div>

        <div class="${ styles.row}">
            <div class="${ styles.column}">
                <div class="${ styles.content}" style="padding-top:5px;padding-bottom:5px;">
                    <p class="${ styles.tripleline}">Data, Digital, Integration, Marketing, Reporting</p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="${ styles.content} ${ styles.empty}">
                    <p></p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="${ styles.content} ${ styles.empty}">
                    <p></p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="${ styles.content} ${ styles.empty}">
                    <p></p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="${ styles.content} ${ styles.empty}">
                    <p></p>
                </div>
            </div>
            <div class="${ styles.column}">
                <div class="${ styles.content} ${ styles.empty}">
                    <p></p>
                </div>
            </div>
        </div>

    

    <div class="${ styles.footercontainer}">
        <div class="${ styles.footeritem}">
            <img src="C:\Users\himan\Downloads\account_balance-24px.svg" />
            <span>Security</span>
        </div>
        <div class="${ styles.footeritem}">
                <img src="C:\Users\himan\Downloads\account_balance-24px.svg" />
                <span>Learning Resources</span>
        </div>        
        <div class="${ styles.footeritem}">
                <img src="C:\Users\himan\Downloads\account_balance-24px.svg" />
                <span>Provide Feedback</span>
        </div>        
        <div class="${ styles.footeritem}">
                <img src="C:\Users\himan\Downloads\account_balance-24px.svg" />
                <span>IT Service Desk</span>
        </div>        
    </div>

</div>`;
  
    this._renderListAsync();
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

  private _renderListAsync(): void {
    // Local environment
    if (Environment.type === EnvironmentType.Local) {
      this._getMockListData().then((response) => {
        this._renderList(response.value);
      });
    }
    else if (Environment.type == EnvironmentType.SharePoint || 
              Environment.type == EnvironmentType.ClassicSharePoint) {
      this._getListData()
        .then((response) => {
          this._renderList(response.value);
        });
    }
  }

  private _getMockListData(): Promise<ISPLists> {
    return MockHttpclient.get()
      .then((data: ISPList[]) => {
        var listData: ISPLists = { value: data };
        return listData;
    }) as Promise<ISPLists>;
  }

  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=Hidden eq false`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
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
