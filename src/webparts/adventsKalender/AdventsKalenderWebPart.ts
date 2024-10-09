import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'AdventsKalenderWebPartStrings';
import AdventsKalender from './components/AdventsKalender';
import { IAdventsKalenderProps } from './components/IAdventsKalenderProps';
import { PropertyFieldDateTimePicker, DateConvention } from '@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker';
import { IDateTimeFieldValue } from "@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker";
import { PropertyFieldFilePicker, IFilePickerResult } from "@pnp/spfx-property-controls/lib/PropertyFieldFilePicker";

export interface IAdventsKalenderWebPartProps {
  description: string;
  FørsteAdventDateTime: IDateTimeFieldValue;  
  AndenAdventDateTime: IDateTimeFieldValue;   
  TredjeAdventDateTime: IDateTimeFieldValue;  
  FjerdeAdventDateTime: IDateTimeFieldValue;
  FørsteAdventURL: string;
  AndenAdventURL: string;
  TredjeAdventURL: string;
  FjerdeAdventURL: string;
  FørsteAdventTitle: string;
  AndenAdventTitle: string;
  TredjeAdventTitle: string;
  FjerdeAdventTitle: string;
  filePickerResult: IFilePickerResult;
  title: string;
}

export default class AdventsKalenderWebPart extends BaseClientSideWebPart<IAdventsKalenderWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const backgroundImageUrl = this.properties.filePickerResult?.fileAbsoluteUrl || '';  // Use the stored URL
  
    const element: React.ReactElement<IAdventsKalenderProps> = React.createElement(
      AdventsKalender,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        førsteAdventDateTime: this.properties.FørsteAdventDateTime, 
        andenAdventDateTime: this.properties.AndenAdventDateTime, 
        tredjeAdventDateTime: this.properties.TredjeAdventDateTime,
        fjerdeAdventDateTime: this.properties.FjerdeAdventDateTime,
        førsteAdventUrl: this.properties.FørsteAdventURL,
        andenAdventUrl: this.properties.AndenAdventURL,
        tredjeAdventUrl: this.properties.TredjeAdventURL,
        fjerdeAdventUrl: this.properties.FjerdeAdventURL,
        backgroundImageUrl: backgroundImageUrl,
        title: this.properties.title,
        førsteAdventTitle: this.properties.FørsteAdventTitle,
        andenAdventTitle: this.properties.AndenAdventTitle,
        tredjeAdventTitle: this.properties.TredjeAdventTitle,
        fjerdeAdventTitle: this.properties.FjerdeAdventTitle
      }
    );
  
    ReactDom.render(element, this.domElement);
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
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
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

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Indstillinger"
          },
          groups: [
            {
              groupName: "Generelt",
              groupFields: [
                PropertyFieldFilePicker('filePicker', {
                  context: this.context,
                  filePickerResult: this.properties.filePickerResult,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onSave: (filePickerResult: IFilePickerResult) => {
                    console.log(filePickerResult);
                    this.properties.filePickerResult = filePickerResult;
                    this.context.propertyPane.refresh();
                    this.render();
                  },
                  onChanged: (filePickerResult: IFilePickerResult) => {
                    console.log(filePickerResult);
                    this.properties.filePickerResult = filePickerResult;
                  },
                  key: "filePickerId",
                  buttonLabel: "Select Image",
                  label: "Background Image",
                  accepts: [".jpg", ".jpeg", ".png"],
                }),
                PropertyPaneTextField('title', {
                  label: 'Page Title',
                  value: this.properties.title
                })
              ]
            },
            {
              groupName: "Første Advent",
              groupFields: [
                PropertyFieldDateTimePicker('FørsteAdventDateTime', {
                  label: 'Select the date',
                  initialDate: this.properties.FørsteAdventDateTime,
                  dateConvention: DateConvention.Date,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  deferredValidationTime: 0,
                  key: 'dateFieldId1',
                  showLabels: false
                }),
                PropertyPaneTextField('FørsteAdventURL', {
                  label: 'Url'
                }),
                PropertyPaneTextField('FørsteAdventTitle', {
                  label: 'Title'
                })
              ]
            },
            {
              groupName: "Anden Advent",
              groupFields: [
                PropertyFieldDateTimePicker('AndenAdventDateTime', {
                  label: 'Select the date',
                  initialDate: this.properties.AndenAdventDateTime,
                  dateConvention: DateConvention.Date,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  deferredValidationTime: 0,
                  key: 'dateFieldId2',
                  showLabels: false
                }),
                PropertyPaneTextField('AndenAdventURL', {
                  label: 'Url'
                }),
                PropertyPaneTextField('AndenAdventTitle', {
                  label: 'Title'
                }),
              ]
            },
            {
              groupName: "Tredje Advent",
              groupFields: [
                PropertyFieldDateTimePicker('TredjeAdventDateTime', {
                  label: 'Select the date',
                  initialDate: this.properties.TredjeAdventDateTime,
                  dateConvention: DateConvention.Date,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  deferredValidationTime: 0,
                  key: 'dateFieldId3',
                  showLabels: false
                }),
                PropertyPaneTextField('TredjeAdventURL', {
                  label: 'Url'
                }),
                PropertyPaneTextField('TredjeAdventTitle', {
                  label: 'Title'
                }),
              ]
            },
            {
              groupName: "Fjerde Advent",
              groupFields: [
                PropertyFieldDateTimePicker('FjerdeAdventDateTime', {
                  label: 'Select the date',
                  initialDate: this.properties.FjerdeAdventDateTime,
                  dateConvention: DateConvention.Date,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  deferredValidationTime: 0,
                  key: 'dateFieldId4',
                  showLabels: false
                }),
                PropertyPaneTextField('FjerdeAdventURL', {
                  label: 'Url'
                }),
                PropertyPaneTextField('FjerdeAdventTitle', {
                  label: 'Title'
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}  