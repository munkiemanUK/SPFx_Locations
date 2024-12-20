import { WebPartContext } from "@microsoft/sp-webpart-base";
import { AppMode } from "../LocationsWebPart";
export interface ILocationsProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  webURL: string;
  context: WebPartContext;
  appMode: AppMode
}