import * as React from "react";
import styles from "./TabLayout.module.scss";
import { ITabLayoutProps } from "./ITabLayoutProps";
import { escape } from "@microsoft/sp-lodash-subset";
import {
  Pivot,
  PivotItem,
  PivotLinkFormat,
  PivotLinkSize
} from "office-ui-fabric-react/lib/Pivot";
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { ITab } from "./ITab";
import * as strings from "TabLayoutWebPartStrings";
import { DisplayMode } from "@microsoft/sp-core-library";
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
export interface ITabLayoutState {
  selectedTab?: string;
}
export default class TabLayout extends React.Component<
  ITabLayoutProps,
  ITabLayoutState
> {
  private _container: HTMLElement;
  private _parents: Map<string, Element>;

  public constructor(props: ITabLayoutProps) {
    super(props);

    this.state = {
      selectedTab:
        props.tabs !== undefined && props.tabs.length > 0
          ? props.tabs[0].name
          : undefined
    };
  }
  public render(): React.ReactElement<ITabLayoutProps> {
    const tabNames = new Array<string>();
    if (this.props.tabs !== undefined) {
      this.props.tabs.forEach((tab: ITab) => {
        if (tabNames.indexOf(tab.name) == -1) {
          tabNames.push(tab.name);
        }
      });
    }

    return (
      <div
        className={styles.tabLayout}
        data-instanceId={this.props.instanceId}
        ref={container => (this._container = container!)}
      >
        <WebPartTitle
          displayMode={this.props.displayMode}
          title={this.props.title}
          updateProperty={this.props.updateTitle}
        />
        {(this.props.tabs == undefined || this.props.tabs.length == 0) && (
          <Placeholder
            iconName="BuildQueueNew"
            iconText={strings.Placeholder_Header}
            description={strings.Placeholder_Description}
            buttonLabel={strings.Placeholder_ButtonLabel}
            onConfigure={this.props.configure}
          />
        )}
        {this.props.tabs !== undefined && this.props.tabs.length > 0 && (
          <div>
            <Pivot
              selectedKey={this.state.selectedTab}
              headersOnly={true}
              getTabId={this.getTabId}
              onLinkClick={this.onTabClick}
              linkFormat={
                this.props.showAsLinks
                  ? PivotLinkFormat.links
                  : PivotLinkFormat.tabs
              }
              linkSize={
                this.props.normalSize
                  ? PivotLinkSize.normal
                  : PivotLinkSize.large
              }
            >
              {tabNames.map((tabName: string) => {
                return (
                  <PivotItem linkText={tabName} itemKey={tabName}></PivotItem>
                );
              })}
            </Pivot>
            {this.props.tabs.map((tab: ITab) => {
              return (
                <div
                  data-htTabName={tab.name}
                  data-htSectionId={tab.sectionId}
                  className={
                    tab.name == this.state.selectedTab ? "" : styles.hidden
                  }
                  aria-labelledby={this.getTabId(tab.name)}
                />
              );
            })}
          </div>
        )}
        {this.props.tabs !== undefined &&
          this.props.tabs.length > 0 &&
          this.props.displayMode == DisplayMode.Edit && (
            <span className={styles.instructions}>
              {strings.EditInstructions}
            </span>
          )}
      </div>
    );
  }
  public componentDidMount() {
    if (this.props.displayMode == DisplayMode.Read) {
      this.moveSections();
    }
  }
  public componentWillUpdate(newProps: ITabLayoutProps) {
    if (
      this.props.displayMode !== newProps.displayMode &&
      newProps.displayMode == DisplayMode.Read
    ) {
      // Ensure our selected state still matches what's available in the tabs
      if (newProps.tabs !== undefined && newProps.tabs.length > 0) {
        if (this.state.selectedTab !== newProps.tabs[0].name) {
          this.setState({
            selectedTab: newProps.tabs[0].name
          });
        }
      } else {
        this.setState({
          selectedTab: undefined
        });
      }
    }
  }
  public componentDidUpdate(prevProps: ITabLayoutProps) {
    if (this.props.displayMode !== prevProps.displayMode) {
      if (this.props.displayMode == DisplayMode.Read) {
        this.moveSections();
      } else {
        this.restoreSections();
      }
    }
  }
  private moveSections = (): void => {
    this._parents = new Map<string, Element>();
    this.props.tabs.forEach((tab: ITab) => {
      const source = document.querySelector(
        `[data-sp-a11y-id="${tab.sectionId}"]`
      );
      const dest = this._container.querySelector(
        `[data-htSectionId="${tab.sectionId}"]`
      );
      if (source && dest) {
        this._parents.set(tab.sectionId, source.parentElement);
        dest.appendChild(source);
      }
    });
  };
  private restoreSections = (): void => {
    this._parents.forEach((parent: Element, sectionId: string) => {
      const tabContent = this._container.querySelector(
        `[data-htSectionId="${sectionId}"]`
      ).firstElementChild;
      if (parent && tabContent) {
        parent.appendChild(tabContent);
      }
    });
  };
  private getTabId = (tabName: string): string => {
    if (this.props !== undefined) {
      return `${this.props.instanceId}_${tabName}`;
    } else {
      console.log(`${tabName} ----- ${this}`);
      return "";
    }
  };
  private onTabClick = (item: PivotItem): void => {
    this.setState({
      selectedTab: item.props.itemKey
    });
    // tslint:disable-next-line: semicolon
  };
}
