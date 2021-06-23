import * as React from 'react';
import * as ReactDom from 'react-dom';
import { SPHttpClient} from "@microsoft/sp-http";
import { IPropertyFieldList, PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as strings from 'DynamicTabsWebPartStrings';
import DynamicTabs from './components/DynamicTabs';
import { IDynamicTabsProps } from './components/IDynamicTabsProps';
import { PropertyPaneDescription } from 'DynamicTabsWebPartStrings';

export interface IDynamicTabsWebPartProps {
  description: string;
  title:string;
  listName: IPropertyFieldList;
  siteurl: string;
  spHttpClient: SPHttpClient;
  context:WebPartContext;
  termsetNameOrID: string;
  fieldName: string;
}

export interface IPropertyControlsTestWebPartProps {
  lists: string | string[]; // Stores the list ID(s)
}

export default class DynamicTabsWebPart extends BaseClientSideWebPart<IDynamicTabsWebPartProps> {

  public render(): void {

    const element: React.ReactElement<IDynamicTabsProps> = React.createElement(
      DynamicTabs,
      {
        description: this.properties.description,
        title: this.properties.title,
        listName: this.properties.listName,
        termsetNameOrID: this.properties.termsetNameOrID,
        fieldName: this.properties.fieldName,
        siteurl: this.context.pageContext.web.absoluteUrl,
        spHttpClient: this.context.spHttpClient,
        context: this.context
      }
    );
    ReactDom.render(element, this.domElement);

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Developed by: Armin Razic 2021"
          },
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('title', {
                  label: 'Enter Title:',
                  maxLength:40
                }),
                PropertyFieldListPicker('listName', {
                  label: 'Select your Document library',
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  includeListTitleAndUrl:true,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId'
                }),
                PropertyPaneTextField("termsetNameOrID", {
                  label: strings.termsetNameOrIDFieldLabel,
                  maxLength:40,
                  description:"Make sure you have created a new column of the 'Managed-Metadata' type that should be named 'Tags'. To create one, go to your Document library settings, and create a new column of the 'Managed-Metadata' type that should be named 'Tags'. For the column settings, make sure you have checked 'Allow multiple values'. For the 'Term Set Settings' make sure you have selected your term set in the terms set tree."
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
