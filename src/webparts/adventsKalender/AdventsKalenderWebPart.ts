import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'AdventsKalenderWebPartStrings';
import AdventsCalendar from './components/AdventsKalender';
import { IAdventsCalendarProps } from './interface/IAdventsKalenderProps';
import { PropertyFieldDateTimePicker, DateConvention } from '@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker';
import { IItemsCalendarWebPartProps } from './interface/IItemCalenderProps';
import { PropertyFieldFilePicker, IFilePickerResult } from "@pnp/spfx-property-controls/lib/PropertyFieldFilePicker";
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import { PropertyFieldPeoplePicker, PrincipalType } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';

export default class AdventsCalendarWebPart extends BaseClientSideWebPart<IItemsCalendarWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const backgroundImageUrl = this.properties.backgroundImageResult?.fileAbsoluteUrl || '';

    const element: React.ReactElement<IAdventsCalendarProps> = React.createElement(
      AdventsCalendar,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        firstItemDateTime: this.properties.firstItemDateTime, 
        secondItemDateTime: this.properties.secondItemDateTime, 
        thirdItemDateTime: this.properties.thirdItemDateTime,
        fourthItemDateTime: this.properties.fourthItemDateTime,
        firstItemUrl: this.properties.firstItemURL,
        secondItemUrl: this.properties.secondItemURL,
        thirdItemUrl: this.properties.thirdItemURL,
        fourthItemUrl: this.properties.fourthItemURL,
        backgroundImageUrl: backgroundImageUrl,
        title: this.properties.title,
        firstItemTitle: this.properties.firstItemTitle,
        secondItemTitle: this.properties.secondItemTitle,
        thirdItemTitle: this.properties.thirdItemTitle,
        fourthItemTitle: this.properties.fourthItemTitle,
        isOneElement: this.properties.isOneElement,
        firstItemImage: this.properties.firstItemImage,
        secondItemImage: this.properties.secondItemImage,
        thirdItemImage: this.properties.thirdItemImage,
        fourthItemImage: this.properties.fourthItemImage,
        textColor: this.properties.textColor,
        group: this.properties.group,
        context: this.context
      }
    );
  
    ReactDom.render(element, this.domElement);
  }
  
  protected async onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
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


protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
  if (propertyPath === 'textColor' && newValue !== oldValue) {
    this.properties.textColor = newValue;
    this.render();
  }
  
  super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
}


  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: strings.GeneralGroupName,
              isCollapsed: true,
              groupFields: [
                PropertyFieldPeoplePicker('group', {
                  label: strings.peoplePickerLabel,
                  initialData: this.properties.group,
                  allowDuplicate: false,
                  principalType: [PrincipalType.SharePoint],
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  context: this.context,
                  properties: this.properties,
                  deferredValidationTime: 0,
                  key: 'peopleFieldId'
                }),
                PropertyPaneToggle('isOneElement', {
                  label: strings.ToggleLabel,
                }),
                PropertyPaneTextField('title', {
                  label: strings.PageLabel,
                  value: this.properties.title
                }),
                PropertyFieldColorPicker('textColor', {
                  label: strings.colorPickerLabel,
                  selectedColor: this.properties.textColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  disabled: false,
                  isHidden: false,
                  alphaSliderHidden: true,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'Precipitation',
                  key: 'colorFieldId'
                })
              ]
            },
            {
              groupName: strings.FirstGroupName,
              isCollapsed: true,
              groupFields: [
                PropertyFieldDateTimePicker('firstItemDateTime', {
                  label: strings.FirstGroupDateLabel,
                  initialDate: this.properties.firstItemDateTime,
                  dateConvention: DateConvention.Date,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  deferredValidationTime: 0,
                  key: 'dateFieldId1',
                  showLabels: false
                }),
                PropertyPaneTextField('firstItemURL', {
                  label: strings.FirstGroupUrlLabel
                }),
                PropertyPaneTextField('firstItemTitle', {
                  label: strings.FirstGroupTitle
                }),
                PropertyFieldFilePicker('firstItemImage', {
                  context: this.context,
                  filePickerResult: this.properties.firstItemImage,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onSave: (filePickerResult: IFilePickerResult) => {
                    console.log(filePickerResult);
                    this.properties.firstItemImage = filePickerResult;
                    this.context.propertyPane.refresh();
                    this.render();
                  },
                  onChanged: (filePickerResult: IFilePickerResult) => {
                    console.log(filePickerResult);
                    this.properties.firstItemImage = filePickerResult;
                  },
                  key: "filePickerId1",
                  buttonLabel: strings.FirstGroupButtonLabel,
                  label: strings.FirstGroupImageLabel,
                  accepts: [".jpg", ".jpeg", ".png"],
                }),
              ]
            },
            {
              groupName: strings.SecondGroupName,
              isCollapsed: true,
              groupFields: [
                PropertyFieldDateTimePicker('secondItemDateTime', {
                  label: strings.SecondGroupDateLabel,
                  initialDate: this.properties.secondItemDateTime,
                  dateConvention: DateConvention.Date,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  deferredValidationTime: 0,
                  key: 'dateFieldId2',
                  showLabels: false
                }),
                PropertyPaneTextField('secondItemURL', {
                  label: strings.SecondGroupUrlLabel
                }),
                PropertyPaneTextField('secondItemTitle', {
                  label: strings.SecondGroupTitle
                }),
                PropertyFieldFilePicker('secondItemImage', {
                  context: this.context,
                  filePickerResult: this.properties.secondItemImage,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onSave: (filePickerResult: IFilePickerResult) => {
                    console.log(filePickerResult);
                    this.properties.secondItemImage = filePickerResult;
                    this.context.propertyPane.refresh();
                    this.render();
                  },
                  onChanged: (filePickerResult: IFilePickerResult) => {
                    console.log(filePickerResult);
                    this.properties.secondItemImage = filePickerResult;
                  },
                  key: "filePickerId2",
                  buttonLabel: strings.SecondGroupButtonLabel,
                  label: strings.SecondGroupImageLabel,
                  accepts: [".jpg", ".jpeg", ".png"],
                }),
              ]
            },
            {
              groupName: strings.ThirdGroupName,
              isCollapsed: true,
              groupFields: [
                PropertyFieldDateTimePicker('thirdItemDateTime', {
                  label: strings.ThirdGroupDateLabel,
                  initialDate: this.properties.thirdItemDateTime,
                  dateConvention: DateConvention.Date,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  deferredValidationTime: 0,
                  key: 'dateFieldId3',
                  showLabels: false
                }),
                PropertyPaneTextField('thirdItemURL', {
                  label: strings.ThirdGroupUrlLabel
                }),
                PropertyPaneTextField('thirdItemTitle', {
                  label: strings.ThirdGroupTitle
                }),
                PropertyFieldFilePicker('thirdItemImage', {
                  context: this.context,
                  filePickerResult: this.properties.thirdItemImage,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onSave: (filePickerResult: IFilePickerResult) => {
                    console.log(filePickerResult);
                    this.properties.thirdItemImage = filePickerResult;
                    this.context.propertyPane.refresh();
                    this.render();
                  },
                  onChanged: (filePickerResult: IFilePickerResult) => {
                    console.log(filePickerResult);
                    this.properties.thirdItemImage = filePickerResult;
                  },
                  key: "filePickerId3",
                  buttonLabel: strings.ThirdGroupButtonLabel,
                  label: strings.ThirdGroupImageLabel,
                  accepts: [".jpg", ".jpeg", ".png"],
                }),
              ]
            },
            {
              groupName: strings.FourthGroupName,
              isCollapsed: true,
              groupFields: [
                PropertyFieldDateTimePicker('fourthItemDateTime', {
                  label: strings.FourthGroupDateLabel,
                  initialDate: this.properties.fourthItemDateTime,
                  dateConvention: DateConvention.Date,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  deferredValidationTime: 0,
                  key: 'dateFieldId4',
                  showLabels: false
                }),
                PropertyPaneTextField('fourthItemURL', {
                  label: strings.FourthGroupUrlLabel
                }),
                PropertyPaneTextField('fourthItemTitle', {
                  label: strings.FourthGroupTitle
                }),
                PropertyFieldFilePicker('fourthItemImage', {
                  context: this.context,
                  filePickerResult: this.properties.fourthItemImage,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onSave: (filePickerResult: IFilePickerResult) => {
                    console.log(filePickerResult);
                    this.properties.fourthItemImage = filePickerResult;
                    this.context.propertyPane.refresh();
                    this.render();
                  },
                  onChanged: (filePickerResult: IFilePickerResult) => {
                    console.log(filePickerResult);
                    this.properties.fourthItemImage = filePickerResult;
                  },
                  key: "filePickerId4",
                  buttonLabel: strings.FourthGroupButtonLabel,
                  label: strings.FourthGroupImageLabel,
                  accepts: [".jpg", ".jpeg", ".png"],
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}