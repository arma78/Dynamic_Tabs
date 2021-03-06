import {IListService } from "../Services/IListService";
import {IVersionService } from "../Services/IVersionService";
import { IUser } from "../Services/IUser";
export interface IDynamicTabsState {
  items:IListService[];
  versionsitems:IVersionService[];
  tabsGrouping:any[];
  seletedTabFilter:string;
  seletedTabFilter2:any;
  hideDialog: boolean;
  hideDialog2: boolean;
  messageFileType:string;
  messageMonth:string;
  messageQuarter:string;
  errormsg:string;
  errormsg2:string;
  errormsg3:string;
  errorstatus:number;
  loadinggear:string;
  loadinggear2:string;
  AuthorValidation:string;
  AuthorButton:string;
  Taxonomyval:string;
  ContainerRendom:string;
  TabRendom:string;
  FilesCounterMesage:string;
  tags: any;
  showhidetaxonomy:string;
  hideforVisitors:string;
  showKey:boolean;
  authtoupdate:string;
  userAuthor:boolean;
  multiSigletaxonomy:boolean;
  DisplayNm:any;
  user:IUser[];
}
