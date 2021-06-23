import {  IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
export interface IListService {
  Id: string;
  FileLeafRef: string;
  Title: string;
  Created:Date;
  File_x0020_Type:string;
  File:any;
  CheckOutType:string;
  CheckedOutByUser:string;
  CheckedOutByUserName:string;
  Tags:any[];
  Created_x0020_By:string;
  User:any;
  Author:any;
}
