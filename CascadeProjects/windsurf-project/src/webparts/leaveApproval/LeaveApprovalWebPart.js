import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'LeaveApprovalWebPartStrings';
import LeaveApproval from './components/LeaveApproval.js';

export default class LeaveApprovalWebPart extends BaseClientSideWebPart {
  constructor() {
    super();
    this._isDarkTheme = false;
    this._environmentMessage = '';
  }

  render() {
    const element = React.createElement(
      LeaveApproval,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
        spHttpClient: this.context.spHttpClient,
        listTitle: this.properties.listTitle || 'LeaveRequests'
      }
    );

    ReactDom.render(element, this.domElement);
  }

  async onInit() {
    if (!this.properties.listTitle) {
      this.properties.listTitle = 'LeaveRequests';
    }
    this._environmentMessage = await this._getEnvironmentMessage();
  }

  async _getEnvironmentMessage() {
    if (!!this.context.sdks.microsoftTeams) {
      const context = await this.context.sdks.microsoftTeams.teamsJs.app.getContext();
      let environmentMessage = '';
      switch (context.app.host.name) {
        case 'Office':
          environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
          break;
        case 'Outlook':
          environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
          break;
        case 'Teams':
        case 'TeamsModern':
          environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
          break;
        default:
          environmentMessage = strings.UnknownEnvironment;
      }
      return environmentMessage;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  onThemeChanged(currentTheme) {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  onDispose() {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  get dataVersion() {
    return Version.parse('1.0');
  }

  getPropertyPaneConfiguration() {
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
                PropertyPaneTextField('listTitle', {
                  label: 'Leave list title'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
