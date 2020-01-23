import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneDropdown
} from "@microsoft/sp-property-pane";

import * as strings from "DirectoryWebPartStrings";
import Directory from "./components/Directory";
import { IDirectoryProps } from "./components/IDirectoryProps";
import {
  IDropdownOption
} from "office-ui-fabric-react";

const orderOptions: IDropdownOption[] = [
  { key: "FirstName", text: strings.FirstName },
  { key: "LastName", text: strings.LastName },
  { key: "Department", text: strings.Department },
  { key: "Location", text: strings.Location },
  { key: "JobTitle", text: strings.JobTitle }
];

export interface IDirectoryWebPartProps {
  title: string;
  searchFirstName: boolean;
  showSort: boolean;
  defaultSort:string;
}

export default class DirectoryWebPart extends BaseClientSideWebPart<IDirectoryWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IDirectoryProps> = React.createElement(
      Directory,
      {
        title: this.properties.title,
        context: this.context,
        searchFirstName: this.properties.searchFirstName,
        defaultSort: this.properties.defaultSort,
        showSort: this.properties.showSort,
        displayMode: this.displayMode,
        updateProperty: (value: string) => {
          this.properties.title = value;
        }
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

    let templateChoice:any;

    if (this.properties.showSort) {
      templateChoice = "";
    }
    else {
      templateChoice = PropertyPaneDropdown("defaultSort", {
        label: strings.defaultSortLabel,
        options: orderOptions,
        selectedKey: "LastName"
    });
    }
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("title", {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneToggle("searchFirstName", {
                  checked: false,
                  label: strings.SearchFirstNameLabel
                }),
                PropertyPaneToggle("showSort", {
                  checked: false,
                  label: strings.ShowSortLabel
                }),
                templateChoice
              ]
            }
          ]
        }
      ]
    };
  }
}
