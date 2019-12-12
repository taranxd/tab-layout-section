import { DisplayMode } from "@microsoft/sp-core-library";
import { ITab } from "./ITab";

export interface ITabLayoutProps {
  instanceId: string;
  title: string;
  displayMode: DisplayMode;
  updateTitle: (title: string) => void;
  configure: () => void;
  tabs: Array<ITab>;
  showAsLinks: boolean;
  normalSize: boolean;
}
