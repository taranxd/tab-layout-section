import { ITab } from "./components/ITab";

export interface ITabLayoutWebPartProps {
  title: string;
  tabs: Array<ITab>;
  showAsLinks: boolean;
  normalSize: boolean;
}
