import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IDateTimeFieldValue, IFilePickerResult, IPropertyFieldGroupOrPerson } from "@pnp/spfx-property-controls";

export interface IAdventsCalendarProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  firstItemDateTime: IDateTimeFieldValue;
  secondItemDateTime: IDateTimeFieldValue;
  thirdItemDateTime: IDateTimeFieldValue;
  fourthItemDateTime: IDateTimeFieldValue;
  firstItemUrl: string;
  secondItemUrl: string;
  thirdItemUrl: string;
  fourthItemUrl: string;
  backgroundImageUrl: string;
  title: string;
  firstItemTitle: string;
  secondItemTitle: string;
  thirdItemTitle: string;
  fourthItemTitle: string;
  isOneElement: boolean;
  firstItemImage: IFilePickerResult;
  secondItemImage: IFilePickerResult;
  thirdItemImage: IFilePickerResult;
  fourthItemImage: IFilePickerResult;
  textColor: string;
  group: IPropertyFieldGroupOrPerson[];
  context: WebPartContext;
}