//Created by: Armin Razic - October 2020
import  {UserProfile, Web } from "sp-pnp-js";
import * as React from 'react';
import { v4 as uuidv4 } from 'uuid';
import styles from './DynamicTabs.module.scss';
import { IDynamicTabsProps } from './IDynamicTabsProps';
import { IDynamicTabsState } from './IDynamicTabsState';
import * as moment from 'moment';
import { IListService } from '../Services/IListService';
import { IUser } from '../Services/IUser';
import { IVersionService } from '../Services/IVersionService';
import { Dialog, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton} from 'office-ui-fabric-react/lib/Button';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { TaxonomyPicker, IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
import  {CheckinType, sp} from "@pnp/sp";
const spinner1: any = require('./assets/small.gif');
import pnp from "sp-pnp-js";
export default class DynamicTabs extends React.Component<IDynamicTabsProps, IDynamicTabsState> {


  constructor(props: IDynamicTabsProps, state: IDynamicTabsState) {
    super(props);
    this._getAuthor = this._getAuthor.bind(this);
    this._versionHistory = this._versionHistory.bind(this);
    this._tabFormSelected =  this._tabFormSelected.bind(this);
    this._tabFormSelected2 =  this._tabFormSelected2.bind(this);
    sp.setup({
      spfxContext: this.props.context
    });
    pnp.setup({
      sp: {
        headers: {
            "Accept": "application/json; odata=verbose"
        }
      }
    });
    this.state = {
    FilesCounterMesage:"none",
    ContainerRendom:"",
    TabRendom:"",
    loadinggear:"none",
    loadinggear2:"none",
    AuthorValidation:"none",
    AuthorButton:"block",
    tags:[],
    errormsg2:"none",
    errormsg3:"none",
    showKey:false,
    items: [],
    user:[],
    Taxonomyval:"none",
    tabsGrouping: [],
    versionsitems:[],
    seletedTabFilter: "",
    seletedTabFilter2:"",
    hideDialog: true,
    hideDialog2:true,
    messageFileType:"none",
    messageMonth:"none",
    userAuthor:false,
    multiSigletaxonomy:false,
    errormsg:"none",
    errorstatus:0,
    DisplayNm:"",
    authtoupdate:"disalowed",
    showhidetaxonomy:"none",
    hideforVisitors:"none",
    messageQuarter:"none"};


  }
  public componentDidMount(): void {

    this.getRandomIDs();

    if(Boolean(this.props.listName) == true && Boolean(this.props.termsetNameOrID) == true)
    {
      this._getDocumentListItems()
      .then((result: Array<IListService>) => {
        this.setState({ items: result });
      });

      this._myValidation2();
      this._myValidation();

    }
    else {
      this.setState({errormsg:"block"});
     }


  }
  public componentDidUpdate(prevProps, prevState) {

    if (this.props.listName !== prevProps.listName
      || this.props.termsetNameOrID !== prevProps.termsetNameOrID
      ) {
        this._myValidation2();
        this._myValidation();
         this._getDocumentListItems()
        .then((result: Array<IListService>) => {
          this.setState({ items: result });

        });

    }
    if (this.state.seletedTabFilter2 !== prevState.seletedTabFilter2) {
      this.setState({FilesCounterMesage:"block" });

      this._getDocumentListItems()
        .then((result: Array<IListService>) => {
          this.setState({ items: result,FilesCounterMesage:"none"});
        });
      }

      if (this.state.seletedTabFilter !== prevState.seletedTabFilter) {
          if(this.state.seletedTabFilter === "")
          {
            this.setState({FilesCounterMesage:"none"});
          }
      }

    if (this.state.items !== prevState.items)
    {

      this._getDocumentListItems()
        .then((result: Array<IListService>) => {
          this.setState({ items: result });
        });

    }

    if (this.state.tags !== prevState.tags) {
      if (this.state.tags.length > 0) {
        this.setState({ showhidetaxonomy: "block" });
      }
      else {
        this.setState({ showhidetaxonomy: "none" });
      }
    }

  }


  private getRandomIDs() {
    var res = uuidv4();
    var res2 = uuidv4();
    this.setState({ ContainerRendom: "Container" + res });
    this.setState({ TabRendom: "Tab" + res2 });
  }

  public render(): React.ReactElement<IDynamicTabsProps> {
    if(Boolean(this.props.listName) == true && this.state.errorstatus == 0)
    {
    return (
      <div className={styles.dynamicTabs}>
        <h3>{this.props.title}</h3>

        <br></br>

        <span style={{ display: this.state.hideforVisitors }}>
         <b>Show Terms:</b>
          <input type="checkbox" onChange={(event) =>this.toggleChange(event)}
           checked={this.state.showKey}/>
         <b>Allow multiple selection for the tag picker:</b>
        <input type="checkbox" onChange={(event) =>this._UpdateMulty(event)}
           checked={this.state.multiSigletaxonomy}/>
        </span>
        <div style={{ display: this._myKeywordsDiv() }} >
        <TaxonomyPicker
        allowMultipleSelections={this.state.multiSigletaxonomy}
        initialValues={this.state.tags}
        termsetNameOrID={this.props.termsetNameOrID}
        panelTitle="Select Document Tags"
        label="Documents Tag Picker"
        context={this.props.context}
        onChange={this.onTaxPickerChange.bind(this)}
        isTermSetSelectable={false} />
       </div>
        <br></br>
        <b>Select Filtering Type: </b><select name="tabform" id="tabform" onChange={(event) => this._tabFormSelected(event)}>
          <option value="">Select</option>
          <option value="Month File Created">(Month) - Document Created</option>
          <option value="Quarter File Created">(Quarter) - Document Created</option>
          <option value="File Types">Document Types</option>
        </select>
        <b> Select Year: </b> <select name="tabform2" id="tabform2"  onChange={(event) => this._tabFormSelected2(event)}>
          <option value="">Select</option>
          <option value={new Date().getFullYear()}>{new Date().getFullYear()}</option>
          <option value={new Date().getFullYear() - 1}>{new Date().getFullYear() - 1}</option>
          <option value={new Date().getFullYear() - 2}>{new Date().getFullYear() - 2}</option>
          <option value={new Date().getFullYear() - 3}>{new Date().getFullYear() - 3}</option>
          <option value={new Date().getFullYear() - 4}>{new Date().getFullYear() - 4}</option>
        </select>  <b>Click on desired tab below.</b>
        <div>
        <b>{this.state.items.length > 0 && this.state.seletedTabFilter2 !== "" && this.state.seletedTabFilter !== "" ? 'We found ' + this.state.items.length + ' documents that were created in year ' + this.state.seletedTabFilter2  : ''}</b>
        <b>{this.state.items.length == 0 && this.state.seletedTabFilter2 !== "" && this.state.seletedTabFilter !== "" ? 'No Documents in your library were created in ' + this.state.seletedTabFilter2 + ' with the following extensions: .pptx, .docx, .pdf or .xlsx!' : ''}</b>
          <img src={require('../components/assets/small.gif')} style={{ display: this.state.FilesCounterMesage, maxHeight:"38px"}} alt="" />
        </div>
        <hr></hr>
       <ul className={styles.tab} id={this.state.TabRendom}>
          {this.state.tabsGrouping.length > 0 && this.state.tabsGrouping.map((TabDisplay, index) => {
            return (
              <li key={index} id={TabDisplay} style={{ background: '#333333' }} ><a href="#" onClick={(event) => this.tabClick(event, TabDisplay)}>{TabDisplay}</a></li>
            );
          })}
        </ul>

        <div id={this.state.ContainerRendom}>
          {this.state.seletedTabFilter !== "" && this.state.seletedTabFilter2 !== "" && this.state.items.length > 0 ? this.state.items.map((listItem, index) => {
            var formatedCreated;
            var formattedGeneral = moment(listItem.Created).format("DD/MMM/YYYY HH:MM").toString();
            var Quarterformatted;
            var kwords = "";

            if(listItem.Tags.length > 0)
            {
              for (let ind = 0; ind < listItem.Tags.length; ind++) {

               kwords += listItem.Tags[ind].Label + " , ";
              }
            }
            else {
              kwords ="Document does not have any assigned tags.";
            }
            if (this.state.seletedTabFilter === "Quarter File Created") {

              Quarterformatted = moment(listItem.Created).quarter().toString();
                return <div key={index} id={'Quarter ' + Quarterformatted} className={styles.tabcontent} style={{ display: 'none' }}>
                  <p><b>Assigned Terms: </b>{kwords}</p>
                  <p><b>Document Name: </b>
                    <b className={styles.Rm}
                      onClick={(event) => this._redirectToPage(event, listItem.FileLeafRef)}>
                      {listItem.FileLeafRef}</b></p>
                  <p><b>Date Created:</b> {formattedGeneral}</p>
                  <p><b>File Type:</b> {listItem.File_x0020_Type}</p>
                  <p><b>Checked Out by User Name:</b>{listItem.CheckedOutByUserName}</p>
                  <p><b>Checked Out by User Email:</b>{listItem.CheckedOutByUser}</p>
                  <p className={styles.pstyle}>
                  <img src={require('../components/assets/small.gif')} style={{ display: this.state.loadinggear2,maxHeight:"38px" }} alt="" />
                  <button className={styles.verhistoryclass} style={{ display: this.state.AuthorButton }}
                    onClick={(event) => this._getAuthor(event, listItem.Author)}>Doc. Author Info</button>
                  <b> | </b>
                  <button className={styles.verhistoryclass} id={listItem.FileLeafRef}
                    onClick={(event) => this._versionHistory(event, listItem.FileLeafRef, listItem.Id)}>Show Version History</button>
                 </p>
                  <p style={{ display: this.state.showhidetaxonomy }}><b style={{ display: this.state.hideforVisitors }}>Assign selected terms to your document: </b>
                  <img src={require('../components/assets/small.gif')} style={{ display: this.state.loadinggear,maxHeight:"38px" }} alt="" />
                  <button style={{ display: this.state.hideforVisitors }} className={styles.verhistoryclass}
                    data-automation-id="addSelectedTerms"
                    title="Assign Selected Terms"
                    onClick={this.updateMultiMeta.bind(this,listItem.Id,listItem.Created_x0020_By,listItem.Tags)}>
                    Assign
                  </button></p>
                </div>;
            }
            else if (this.state.seletedTabFilter === "File Types") {
              return <div key={index} id={listItem.File_x0020_Type} className={styles.tabcontent} style={{ display: 'none' }}>
                <p><b>Assigned Terms: </b>{kwords}</p>
                <p><b>Document Name: </b><b className={styles.Rm} onClick={(event) => this._redirectToPage(event, listItem.FileLeafRef)}>{listItem.FileLeafRef}</b></p>
                <p><b>Date Created:</b> {formattedGeneral}</p>
                <p><b>Checked Out by User Name:</b>{listItem.CheckedOutByUserName}</p>
                <p><b>Checked Out by User Email:</b>{listItem.CheckedOutByUser}</p>
                <p><b>File Type:</b> {listItem.File_x0020_Type}</p>

                   <p className={styles.pstyle}>
                  <img src={require('../components/assets/small.gif')} style={{ display: this.state.loadinggear2,maxHeight:"38px" }} alt="" />
                  <button className={styles.verhistoryclass} style={{ display: this.state.AuthorButton }}
                    onClick={(event) => this._getAuthor(event, listItem.Author)}>Doc. Author Info</button>
                  <b>  |  </b>
                  <button className={styles.verhistoryclass} id={listItem.FileLeafRef}
                    onClick={(event) => this._versionHistory(event, listItem.FileLeafRef,listItem.Id)}>Show Version History</button>
                </p>
                 <p style={{ display: this.state.showhidetaxonomy }}><b style={{ display: this.state.hideforVisitors }}>Assign selected terms to your document: </b>
                 <img src={require('../components/assets/small.gif')} style={{ display: this.state.loadinggear,maxHeight:"38px" }} alt="" />
                  <button style={{ display: this.state.hideforVisitors }} className={styles.verhistoryclass}
                    data-automation-id="addSelectedTerms"
                    title="Assign Selected Terms"
                    onClick={this.updateMultiMeta.bind(this,listItem.Id,listItem.Created_x0020_By,listItem.Tags)}>
                    Assign
                  </button></p>
              </div>;
            } else if (this.state.seletedTabFilter === "Month File Created") {
              formatedCreated = moment(listItem.Created).format("MMMM").toString();
              return <div key={index} id={formatedCreated} className={styles.tabcontent} style={{ display: 'none' }}>
                <p><b>Assigned Terms: </b>{kwords}</p>
                <p><b>Document Name: </b><b className={styles.Rm} onClick={(event) => this._redirectToPage(event, listItem.FileLeafRef)}>{listItem.FileLeafRef}</b></p>
                <p><b>Date Created:</b> {formattedGeneral}</p>
                <p><b>Checked Out by User Name:</b>{listItem.CheckedOutByUserName}</p>
                <p><b>Checked Out by User Email:</b>{listItem.CheckedOutByUser}</p>
                <p><b>File Type:</b> {listItem.File_x0020_Type}</p>
                <p className={styles.pstyle}>
                   <img src={require('../components/assets/small.gif')} style={{ display: this.state.loadinggear2,maxHeight:"38px" }} alt="" />
                  <button className={styles.verhistoryclass} style={{ display: this.state.AuthorButton }}
                    onClick={(event) => this._getAuthor(event, listItem.Author)}>Doc. Author Info</button>
                  <b>  |  </b>
                  <button className={styles.verhistoryclass} id={listItem.FileLeafRef}
                    onClick={(event) => this._versionHistory(event, listItem.FileLeafRef, listItem.Id)}>Show Version History</button>
                </p>
               <p style={{ display: this.state.showhidetaxonomy }}><b style={{ display: this.state.hideforVisitors }}>Assign selected terms to your document: </b>
               <img src={require('../components/assets/small.gif')} style={{ display: this.state.loadinggear,maxHeight:"38px" }} alt="" />
                      <button style={{ display: this.state.hideforVisitors }} className={styles.verhistoryclass}
                    data-automation-id="addSelectedTerms"
                    title="Assign Selected Terms"
                    onClick={this.updateMultiMeta.bind(this,listItem.Id,listItem.Created_x0020_By,listItem.Tags)}>
                    Assign
                  </button></p>
              </div>;
              // tslint:disable-next-line:no-unused-expression

            }

          }):<div>
             </div>}
        </div>

        <div><b>{this.state.seletedTabFilter === "" ? 'Please, Select the filtering type from drop-down box' : ''}</b></div>
        <div><b>{ this.state.seletedTabFilter2 === "" ? 'Please, Select the Year from drop-down box' : ''}</b></div>
        <div style={{ display: this.state.messageQuarter }} ><b>No documents in your library were created in this quarter!</b></div>
        <div style={{ display: this.state.messageMonth }} ><b>No documents in your library were created in this month!</b></div>
        <div style={{ display: this.state.messageFileType }} ><b>No documents in your library were created with this extension!</b></div>



        <Dialog isClickableOutsideFocusTrap={false} hidden={this.state.hideDialog} onDismiss={this._closeDialog}>
            <Label><h3>Document Version History</h3></Label>
        <div style={{ color: 'black', height: '250px', overflow: 'auto'}}>
            {this.state.versionsitems.length > 0 ? this.state.versionsitems.map((VerSionDisplay, index) => {
              return (
                <div><p><b>Version: </b>{VerSionDisplay.VersionLabel}</p>
                  <p><b>Version Date: </b>{VerSionDisplay.Created}</p>
                  <p><b>Version Comment: </b>{VerSionDisplay.CheckInComment}</p>
                  <hr></hr></div>
              );
            }) : <div><b>No vesrions created yet for this document!</b></div>}
          </div>
           <DialogFooter>
              <PrimaryButton onClick={this._closeDialog} text="Close" />
            </DialogFooter>
          </Dialog>

          <Dialog isClickableOutsideFocusTrap={false} hidden={this.state.hideDialog2} onDismiss={this._closeDialog2}>
            <Label><h3>Document Author Info</h3></Label>
        <div style={{ color: 'black', height: '250px', overflow: 'auto'}}>
            {this.state.user.length > 0 ? this.state.user.map((UserDisplay, index) => {
              return (

                <div>
                <p><b>User Name: </b>{UserDisplay.Title}</p>
                <p><b>User Email: </b>{UserDisplay.Email}</p>
                <p><b>'Site Visitor' Group Member: </b>{UserDisplay.userVisitor}</p>
                <p><b>'Site Member' Group Member: </b>{UserDisplay.userMember}</p>
                <p><b>'Site Owner' Group Member: </b>{UserDisplay.userOwner}</p>
                <p><b>User ID: </b>{UserDisplay.Id}</p>
                  <hr></hr></div>
              );
            }) : <div><b>No Info for this Author!</b></div>}
          </div>
          <DialogFooter>
              <PrimaryButton onClick={this._closeDialog2} text="Close" />
            </DialogFooter>
          </Dialog>
      </div>

    );
    }
    else {
      return (<div><div style={{ display: this.state.errormsg, color: "red" }} >
        <h2>
          <ul>
            <li>Please, Go to edit mode.</li>
            <li>Select your Document library.</li>
            <li>Enter Term Set Name</li>
            <li>Press "Apply" button.</li>
          </ul>
        </h2>
        </div>
        <b>{this.state.errorstatus == 400 ? 'Please check your library name, Term Set, and TAGS field!' : ''}</b>
        <b>{this.state.errorstatus == 404 ? 'Please check your library name, Term Set, and TAGS field!' : ''}</b>
        <b>{this.state.errorstatus == 500 ? 'Please check your library name, Term Set, and TAGS field!' : ''}</b>

        </div>

        );

     }


  }




      // Handle Toggle Change
      private _UpdateMulty(e)
       {
         this.setState({multiSigletaxonomy:e.target.checked});
       }

  private _closeDialog = (): void => {
    this.setState({ hideDialog: true });
  }
  private _closeDialog2 = (): void => {
    this.setState({ hideDialog2: true });
  }

   // tslint:disable-next-line:typedef
   private toggleChange(e) {
    this.setState({
      showKey: e.target.checked
    });
  }


public onTaxPickerChange(terms : IPickerTerms) {

  this.setState({ tags: terms });
  if(this.state.tags.length > 0)
    {
      this.setState({showhidetaxonomy:"block"});
    }
    else {
      this.setState({showhidetaxonomy:"none"});
    }
}



  public async updateMultiMeta(id:string, CreatedBy:string,tagsSaved:any): Promise<any> {
   this.setState({hideforVisitors:"none",loadinggear:"block"});

    let currUserEmail = "";
    let userAuthor:boolean = false;
                const profile = await sp.profiles.myProperties.get();
      currUserEmail = profile.AccountName;

     if(currUserEmail.toString() === CreatedBy.toString() )
     {
       userAuthor = true;
     }
     else {
      userAuthor = false;
     }
        let termsString: string = '';
        let pickedterms: IPickerTerms = [];
        pickedterms = this.state.tags;
        pickedterms.forEach(term => {
            termsString += `-1;#${term.name}|${this.cleanGuid(term.key)};#`;
        });

    if (this.state.authtoupdate === "alowed" || userAuthor == true)
    {




    let listName: string = this.props.listName.title;
    let fieldName: string = "Tags";
    let itemId: number = +id;
    const data = {};
    const list = sp.web.lists.getByTitle(listName);
    const field = await list.fields.getByTitle(`${fieldName}_0`).get();
      if (this.state.multiSigletaxonomy == false) {
       var kwords = "";
       if(tagsSaved.length > 0)
       {
         // tslint:disable-next-line:no-shadowed-variable
         for (let ind = 0; ind < tagsSaved.length; ind++) {
          kwords += "-1;#" + tagsSaved[ind].Label + "|" + tagsSaved[ind].TermGuid + ";#";
         }
       }
       else {
         kwords ="";
       }
       data[field.InternalName] = kwords +  "-1;#" + pickedterms[0].name + "|" + pickedterms[0].key + ";#";
       sp.web.getFileByServerRelativeUrl(this.props.listName.url).checkin('Discard', 2).then(() => alert("Discarded"))
      .then((response) => { return response; }).catch((e) => { console.log("Discard Failed, either you don't have permission, it's already discarded, or you never had it checked out."); });

       await list.items.getById(itemId).update(data);

      }
      else {
        sp.web.getFileByServerRelativeUrl(this.props.listName.url).checkin('Discard', 2).then(() => alert("Discarded"))
        .then((response) => { return response; }).catch((e) => { console.log("Discard Failed, either you don't have permission, it's already discarded, or you never had it checked out."); });
        data[field.InternalName] = termsString;
        await list.items.getById(itemId).update(data);
      }


  }
  else {
    alert("You are not authorized to assign tags to this document!");
  }
  this.setState({hideforVisitors:"block", loadinggear:"none"});
}

public cleanGuid(guid: string): string {
  if (guid !== undefined) {
      return guid.replace('/Guid(', '').replace('/', '').replace(')', '');
  } else {
      return '';
  }
}


public async _getUserId(): Promise<string> {
  // tslint:disable-next-line:typedef
  let myweb = new Web(this.props.context.pageContext.web.absoluteUrl);
  return await myweb.currentUser.get().then((r: UserProfile) => {
    // tslint:disable-next-line:no-string-literal
    return r["Id"];
  });
}

// tslint:disable-next-line:typedef
public _myValidation() {
  this._AuthorizedToUpdate().then(res => {
       if(res === "bingo") {
        this.setState({authtoupdate:"alowed"});
      } else {
        this.setState({authtoupdate:"disalowed"});
      }
    });
  }

// tslint:disable-next-line:typedef
private async _AuthorizedToUpdate(): Promise<string> {
 let currUserEmail = "";
 var userinOwnerGroup = "";
 const profile = await sp.profiles.myProperties.get();
 currUserEmail = profile.Email;
   const memberGroup = await sp.web.associatedOwnerGroup.get();
   const groupID = memberGroup.Id;
   const usersd = await sp.web.siteGroups.getById(groupID).users.get();
   for (let index = 0; index < usersd.length; index++) {
     if(usersd[index].Email === currUserEmail) {
      userinOwnerGroup = "bingo";
     }
   }
   return  userinOwnerGroup;

}

// tslint:disable-next-line:typedef
public _myValidation2() {
  this.setState({Taxonomyval:"block"});
  this._AuthorizedToSee().then(res => {
       if(res !== "bingo" && this.state.authtoupdate !== "alowed") {
        this.setState({hideforVisitors:"none"});
      } else {
        this.setState({hideforVisitors:"block"});
      }
    });
     this.setState({Taxonomyval:"none"});

  }

// tslint:disable-next-line:typedef
private async _AuthorizedToSee(): Promise<string> {
  let currUserEmail = "";
  var userMember = "";
  const profile = await sp.profiles.myProperties.get();
  currUserEmail = profile.Email;
    const MembersGroup = await sp.web.associatedMemberGroup.get();
    const groupID = MembersGroup.Id;
    const usersd = await sp.web.siteGroups.getById(groupID).users.get();
    for (let index = 0; index < usersd.length; index++) {
      if(usersd[index].Email === currUserEmail) {
        userMember = "bingo";
      }
    }
    return  userMember;
}

 // tslint:disable-next-line:typedef
 private  _myKeywordsDiv() {
  if (this.state.showKey === false) {
    return "none";
  } else if (this.state.showKey === true) {
    return "block";
  }
  return "";
}

public tabClick(evt, paramTabQuery) {
  this.setState({messageQuarter:"none",messageFileType:"none",messageMonth:"none"});

    if(this.state.tags.length > 0)
    {
      this.setState({showhidetaxonomy:"block"});
    }
    else {
      this.setState({showhidetaxonomy:"none"});
    }

    let q:number[] = [];
    let m:string[] = [];
    let ext:string[] = [];
    let found:number[] = [];
    let mfound:string[] = [];
    let extfound:string[] = [];
    var i, tabcontent, tablinks;
        tablinks = document.getElementById(this.state.TabRendom).getElementsByTagName("li");
        for (i = 0; i < tablinks.length; i++) {
          tablinks[i].style.background = "#333333";
        }
        var t = document.getElementById(this.state.TabRendom);
        t.children.namedItem(paramTabQuery).setAttribute("style", "background-color:orange");

    tabcontent = document.getElementById(this.state.ContainerRendom).getElementsByTagName("div");
    for (i = 0; i < tabcontent.length; i++) {
      if (tabcontent[i].id === paramTabQuery) {
        tabcontent[i].style.display = "block";
      }
      else {
        tabcontent[i].style.display = "none";
      }
    }


    if (this.state.items.length > 0) {
      if(paramTabQuery === 'Quarter 4' || paramTabQuery === 'Quarter 3' || paramTabQuery === 'Quarter 2'|| paramTabQuery === 'Quarter 1')
      {
      this.state.items.map((lItem) => {
        q.push(moment(lItem.Created).utc().quarter());
      });
      }
      if(paramTabQuery !== 'Quarter 4' || paramTabQuery !== 'Quarter 3' || paramTabQuery !== 'Quarter 2'|| paramTabQuery !== 'Quarter 1')
      {
      this.state.items.map((mlItem) => {
        m.push(moment(mlItem.Created).format('M'));
      });
      }
      if(paramTabQuery === 'pptx' || paramTabQuery === 'docx' || paramTabQuery === 'pdf'|| paramTabQuery === 'xlsx')
      {
      this.state.items.map((extlItem) => {
        ext.push(extlItem.File_x0020_Type);
      });
      }

      if (paramTabQuery === 'pdf') {
        // tslint:disable-next-line:no-function-expression
        extfound = ext.filter(function (string) {
          return string == 'pdf';
        });
       }
       if (paramTabQuery === 'pptx') {
        // tslint:disable-next-line:no-function-expression
        extfound = ext.filter(function (string) {
          return string == 'pptx';
        });
       }
       if (paramTabQuery === 'docx') {
        // tslint:disable-next-line:no-function-expression
        extfound = ext.filter(function (string) {
          return string == 'docx';
        });
       }
       if (paramTabQuery === 'xlsx') {
        // tslint:disable-next-line:no-function-expression
        extfound = ext.filter(function (string) {
          return string == 'xlsx';
        });
       }

      if (paramTabQuery === 'January') {
        // tslint:disable-next-line:no-function-expression
        mfound = m.filter(function (string) {
          return string == '1';
        });
       }
       if (paramTabQuery === 'February') {
        // tslint:disable-next-line:no-function-expression
        mfound = m.filter(function (string) {
          return string == '2';
        });

       }
       if (paramTabQuery === 'March') {
        // tslint:disable-next-line:no-function-expression
        mfound = m.filter(function (string) {
          return string == '3';
        });

       }
       if (paramTabQuery === 'April') {
        // tslint:disable-next-line:no-function-expression
        mfound = m.filter(function (string) {
          return string == '4';
        });

       }
       if (paramTabQuery === 'May') {
        // tslint:disable-next-line:no-function-expression
        mfound = m.filter(function (string) {
          return string == '5';
        });

       }
       if (paramTabQuery === 'Juny') {
        // tslint:disable-next-line:no-function-expression
        mfound = m.filter(function (string) {
          return string == '6';
        });

       }
       if (paramTabQuery === 'July') {
        // tslint:disable-next-line:no-function-expression
        mfound = m.filter(function (string) {
          return string == '7';
        });

       }
       if (paramTabQuery === 'August') {
        // tslint:disable-next-line:no-function-expression
        mfound = m.filter(function (string) {
          return string == '8';
        });

       }
       if (paramTabQuery === 'September') {
        // tslint:disable-next-line:no-function-expression
        mfound = m.filter(function (string) {
          return string == '9';
        });

       }
       if (paramTabQuery === 'October') {
        // tslint:disable-next-line:no-function-expression
        mfound = m.filter(function (string) {
          return string == '10';
        });

       }
       if (paramTabQuery === 'November') {
        // tslint:disable-next-line:no-function-expression
        mfound = m.filter(function (string) {
          return string == '11';
        });

       }
       if (paramTabQuery === 'December') {
        // tslint:disable-next-line:no-function-expression
        mfound = m.filter(function (string) {
          return string == '12';
        });

       }


        if (paramTabQuery === 'Quarter 1') {
          // tslint:disable-next-line:no-function-expression
          found = q.filter(function (number) {
            return number == 1;
          });
         }
        else if (paramTabQuery === 'Quarter 2') {
          // tslint:disable-next-line:no-function-expression
          found = q.filter(function (number) {
            return number == 2;
          });
        }
        else if (paramTabQuery === 'Quarter 3') {
          // tslint:disable-next-line:no-function-expression
          found = q.filter(function (number) {
            return number == 3;
          });
        }
        else if (paramTabQuery === 'Quarter 4') {
          // tslint:disable-next-line:no-function-expression
          found = q.filter(function (number) {
            return number == 4;
          });
        }

      if (paramTabQuery === 'Quarter 4'
       || paramTabQuery === 'Quarter 3'
       || paramTabQuery === 'Quarter 2'
       || paramTabQuery === 'Quarter 1'
       ) {
        if (found.length == 0) {
          this.setState({ messageQuarter: "block", messageMonth: "none", messageFileType:"none" });
        }
      }
      else if (paramTabQuery === 'January'
       || paramTabQuery === 'February'
       || paramTabQuery === 'March'
       || paramTabQuery === 'April'
       || paramTabQuery === 'May'
       || paramTabQuery === 'June'
       || paramTabQuery === 'July'
       || paramTabQuery === 'August'
       || paramTabQuery === 'September'
       || paramTabQuery === 'October'
       || paramTabQuery === 'November'
       || paramTabQuery === 'December'
       ) {
        if (mfound.length == 0) {
          this.setState({ messageMonth: "block",messageQuarter:"none", messageFileType:"none" });
        }
      }
      else if(
         paramTabQuery === 'pptx'
        || paramTabQuery === 'docx'
        || paramTabQuery === 'pdf'
        || paramTabQuery === 'xlsx')
      {
        if (extfound.length == 0) {
          this.setState({ messageFileType: "block",messageMonth:"none",messageQuarter:"none" });
        }
      }
    }

  }
  private _getDocumentListItems(): Promise<IListService[]> {

    return new Promise<IListService[]>((resolve:(any) => void, reject: (error: any) => void): void => {
    if(this.state.seletedTabFilter2 !== "" && this.state.seletedTabFilter !== "")
      {
      let year = this.state.seletedTabFilter2;
      var SHAREPOINT_LIST: string = this.props.listName.title;


      sp.web.lists.getByTitle(SHAREPOINT_LIST).items
        .select("Id","Created_x0020_By","Tags","Author/Id","FileLeafRef", "Title", "Created", "File_x0020_Type","File/CheckOutType")
        .expand('Author','File', 'File/CheckOutType','File/CheckedOutByUser','FieldValuesAsText')
        .filter((`File_x0020_Type eq  'pptx' or File_x0020_Type eq  'docx' or  File_x0020_Type eq  'pdf'  or  File_x0020_Type eq  'xlsx'`))
        .orderBy('Created', false)
        .get()
        .then((response: any[]) => {

          if (Boolean(response) == false || response.length == 0) {
            this.setState({ errormsg3:"block"});
            resolve(-1);
          }
          else if (Boolean(response) == true || response.length > 0)
          {
          let user:any = "";
          let uemail = '';
          let uname = '';
          let items: IListService[] = [];




          // tslint:disable-next-line:no-function-expression
          response.forEach(function (item: IListService) {

           // tslint:disable-next-line:no-function-expression

            if (moment(item.Created).format('yyyy').toString() === year) {



              if (item.File.CheckOutType.toString() === '0') {
                uemail = item.File.CheckedOutByUser.Email;
                uname = item.File.CheckedOutByUser.Title;
              }
              else {
                uemail = "Item is not checked out.";
                uname = "Item is not checked out.";
              }


              items.push({
                Id:item.Id,
                FileLeafRef: item.FileLeafRef,
                Title: item.Title,
                Created: item.Created,
                File_x0020_Type: item.File_x0020_Type,
                File: item.File,
                CheckOutType: item.File.CheckOutType.toString(),
                CheckedOutByUser: uemail,
                CheckedOutByUserName: uname,
                Tags:item.Tags,
                Created_x0020_By:item.Created_x0020_By,
                User:user,
                Author:item.Author.Id
              });

            }
          });

            this.setState({items: items, errormsg:"none",errorstatus:0, errormsg3:"none"});
            resolve(items);
        }}, (error: any): void => {
          if(error.status == 404 || error.status == 400)
          {
            this.setState({errorstatus:error.status});
          }
          reject(error);
        });

        }
    });
  }





  public _versionHistory(event, itemId, id) {

      this._getDocVersions(itemId, id)
      .then((resultver: Array<IVersionService>) => {
        this.setState({ versionsitems: resultver });
      });

      this.setState({ hideDialog:false });
  }

  public async _getAuthor(e, UserId:number) {
    this.setState({AuthorButton:"none",loadinggear2:"block"});
    const user = await sp.web.siteUsers.getById(UserId).get();
    const MembersGroup1 = await sp.web.associatedVisitorGroup.get();
    const MembersGroup2 = await sp.web.associatedMemberGroup.get();
    const MembersGroup3 = await sp.web.associatedOwnerGroup.get();
    const groupID1 = MembersGroup1.Id;
    const groupID2 = MembersGroup2.Id;
    const groupID3 = MembersGroup3.Id;

    const usersd1 = await sp.web.siteGroups.getById(groupID1).users.get();
    const usersd2 = await sp.web.siteGroups.getById(groupID2).users.get();
    const usersd3 = await sp.web.siteGroups.getById(groupID3).users.get();
    let userVisitor:string = "No";
    let userMember:string = "No";
    let userOwner:string = "No";

    for (let index = 0; index < usersd1.length; index++) {
      if(usersd1[index].Id === user.Id) {
        userVisitor = "Yes";
        break;
      }
    }
    for (let index = 0; index < usersd3.length; index++) {
      if(usersd3[index].Id === user.Id) {
        userOwner = "Yes";
        break;
      }
    }
    for (let index = 0; index < usersd2.length; index++) {
      if(usersd2[index].Id === user.Id) {
        userMember = "Yes";
        break;
      }
    }
        let items: IUser[] = [];
        items.push({
        Id:user.Id,
        Title:user.Title,
        Email:user.Email,
        userVisitor:userVisitor,
        userMember:userMember,
        userOwner:userOwner
        });

    this.setState({ user: items });
    this.setState({AuthorButton:"block",loadinggear2:"none"});
    this.setState({ hideDialog2:false });
}



  private  _redirectToPage(event, itemId) {

    var SHAREPOINT_LIST: string = this.props.listName.url;
    // As we are creating the  nav link to documents library item
    // SP url is using 'shared documents' for doc library name
    try {

       location.href = this.props.siteurl + SHAREPOINT_LIST + "/" + itemId;

    }
    catch (e) {
      console.log(e);
    }
  }

  public _tabFormSelected2(event) {

    this.setState({seletedTabFilter2: event.target.value,
      messageQuarter:"none",
      messageFileType:"none",
      messageMonth:"none",
      });
    var tablinks = document.getElementById(this.state.TabRendom).getElementsByTagName("li");
    for (var i = 0; i < tablinks.length; i++) {
      tablinks[i].style.background = "#333333";
    }
    var tabcontent = document.getElementById(this.state.ContainerRendom).getElementsByTagName("div");
    for (var j = 0; j < tabcontent.length; j++) {
      tabcontent[j].style.display = "none";
    }
  }

  public _tabFormSelected(event) {
    this.setState({messageQuarter:"none",messageFileType:"none",messageMonth:"none"});
    this.setState({ tabsGrouping: [] });
    this.setState({ seletedTabFilter: event.target.value });
    if (event.target.value === "File Types") {
      this.setState({ tabsGrouping: ["pptx", "docx", "pdf", "xlsx"] });
    }
    else if (event.target.value === "Quarter File Created") {
      this.setState({ tabsGrouping: ["Quarter 1", "Quarter 2", "Quarter 3", "Quarter 4"] });
    }
    else {
      this.setState({
        tabsGrouping: ["January", "February", "March", "April", "May", "June",
          "July", "August", "September", "October", "November", "December"]
      });
    }
    //Reset tabs when drop down change
    var tablinks = document.getElementById(this.state.TabRendom).getElementsByTagName("li");
    for (var i = 0; i < tablinks.length; i++) {
      tablinks[i].style.background = "#333333";
    }
    var tabcontent = document.getElementById(this.state.ContainerRendom).getElementsByTagName("div");
    for (var j = 0; j < tabcontent.length; j++) {
      tabcontent[j].style.display = "none";
    }
  }

  private _getDocVersions(docname: string, docId:number): Promise<IVersionService[]> {
    return new Promise<IVersionService[]>((resolve: any) => {
      let endpointUrl = "";

       endpointUrl = this.props.siteurl + "/_api/web/GetFolderByServerRelativeUrl('" + this.props.listName.url +"')/Files('" + docname + "')/versions/";


      fetch(endpointUrl, {
        method: 'GET',
        credentials: 'same-origin',
        headers: {
          'Accept': `application/json; odata=verbose`,
          'Content-Type': `application/json; odata=verbose`,
        },
      }).then(response => response.json())
        .then((data: any) => {
          let veritems: IVersionService[] = [];
          // tslint:disable-next-line:no-function-expression
          data.d.results.forEach(function (item: IVersionService) {
            veritems.push({
              VersionLabel: item.VersionLabel,
              Created: moment(item.Created).format("DD/MMM/YYYY HH:MM").toString(),
              CheckInComment: item.CheckInComment
            });
          });

         /*  sp.web.getList(this.props.listName.url)
          .items.getById(docId)
          .expand('Versions')
          .get()
          .then((dat2:any) => {
            var x:any = dat2.Versions[Versions.length -1];
              veritems.push({
              VersionLabel:"5",
              Created: moment(x.Last_x005f_x0020_x005f_Modified).format("DD/MMM/YYYY HH:MM").toString(),
              CheckInComment: "Last Version created by: " + x.Editor[2].LookupValue
            });
          }); */

          this.setState({ versionsitems: veritems });
          resolve(veritems);
        });
    });



  }





}
