import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneToggle,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'HelloWorldWebPartStrings';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';
import HelloWorld from './components/HelloWorld';
import { IHelloWorldProps } from './components/IHelloWorldProps';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import styles from './components/HelloWorld.module.scss';

export interface IHelloWorldWebPartProps {
  description: string;
  test1: boolean;
  test2: string;
  test3: boolean;
  context: WebPartContext;
}

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=Hidden eq false`,
      SPHttpClient.configurations.v1
    )
    .then((response: SPHttpClientResponse) => response.json())
    .catch(() => ({ value: [] }));
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

    const listContainer = this.domElement.querySelector('#spListContainer');
    if (listContainer) {
      listContainer.innerHTML = html;
    }
  }

  private _renderListAsync(): void {
    this._getListData()
      .then((response) => {
        this._renderList(response.value);
      })
      .catch(() => {});
  }

  public render(): void {
    const element: React.ReactElement<IHelloWorldProps> = React.createElement(
      HelloWorld,
      {
        description: this.properties.description,
        test1: this.properties.test1,
        test2: this.properties.test2,
        test3: this.properties.test3,
        _environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        isDarkTheme: this._isDarkTheme,
        pageContext: this.context.pageContext
      }
    );

    ReactDom.render(element, this.domElement);

    setTimeout(() => this._renderListAsync(), 0);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
    super.onDispose();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: "Description",
                }),
                PropertyPaneCheckbox("test1", {
                  text: "Check this box to enable feature",
                }),
                PropertyPaneDropdown("test2", {
                  label: "Choose an option",
                  options: [
                    { key: "option1", text: "Option 1" },
                    { key: "option2", text: "Option 2" },
                    { key: "option3", text: "Option 3" },
                  ],
                }),
                PropertyPaneToggle("test3", {
                  label: "Enable feature",
                  onText: "Enabled",
                  offText: "Disabled",
                }),
              ],
            },
          ],
        }
      ]
    };
  }
}
