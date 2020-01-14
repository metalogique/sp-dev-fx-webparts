import { WebPartContext } from "@microsoft/sp-webpart-base";
import { DisplayMode } from "@microsoft/sp-core-library";
export interface IDirectoryProps {
  title: string;
  displayMode: DisplayMode;
  context: WebPartContext;
  searchFirstName: boolean;
  showSort: boolean;
  defaultSort: string;
  updateProperty: (value: string) => void;
}
