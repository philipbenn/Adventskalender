import { IDateTimeFieldValue } from "@pnp/spfx-property-controls";

export interface IAdventsKalenderProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  førsteAdventDateTime: IDateTimeFieldValue;
  andenAdventDateTime: IDateTimeFieldValue;
  tredjeAdventDateTime: IDateTimeFieldValue;
  fjerdeAdventDateTime: IDateTimeFieldValue;
  førsteAdventUrl: string;
  andenAdventUrl: string;
  tredjeAdventUrl: string;
  fjerdeAdventUrl: string;
  backgroundImageUrl: string;
  title: string;
  førsteAdventTitle: string;
  andenAdventTitle: string;
  tredjeAdventTitle: string;
  fjerdeAdventTitle: string;
}