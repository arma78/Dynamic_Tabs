import {SPHttpClient} from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IPropertyFieldList } from "@pnp/spfx-property-controls/lib/PropertyFieldListPicker";
export interface IDynamicTabsProps {
  spHttpClient: SPHttpClient;
  listName: IPropertyFieldList;
  description: string;
  title: string;
  siteurl: string;
  context:WebPartContext;
  termsetNameOrID:string;
  fieldName:string;
}
