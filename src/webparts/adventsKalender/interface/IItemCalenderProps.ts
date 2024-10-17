import { IDateTimeFieldValue, IFilePickerResult } from "@pnp/spfx-property-controls";
import { IPropertyFieldGroupOrPerson } from "@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker";

export interface IItemsCalendarWebPartProps {
    description: string;
    firstItemDateTime: IDateTimeFieldValue;  
    secondItemDateTime: IDateTimeFieldValue;   
    thirdItemDateTime: IDateTimeFieldValue;  
    fourthItemDateTime: IDateTimeFieldValue;
    firstItemURL: string;
    secondItemURL: string;
    thirdItemURL: string;
    fourthItemURL: string;
    firstItemTitle: string;
    secondItemTitle: string;
    thirdItemTitle: string;
    fourthItemTitle: string;
    backgroundImageResult: IFilePickerResult;
    title: string;
    isOneElement: boolean;
    firstItemImage: IFilePickerResult;
    secondItemImage: IFilePickerResult;
    thirdItemImage: IFilePickerResult;
    fourthItemImage: IFilePickerResult;
    textColor: string;
    group: IPropertyFieldGroupOrPerson[];
}