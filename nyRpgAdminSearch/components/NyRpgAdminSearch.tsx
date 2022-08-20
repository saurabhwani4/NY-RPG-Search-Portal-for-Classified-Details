import * as React from 'react';
import { INyRpgAdminSearchProps } from './INyRpgAdminSearchProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as $ from 'jquery';
require('../../../../node_modules/bootstrap/dist/css/bootstrap.min.css');
import '../styles.css';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import Dropdown from 'react-dropdown';
import 'react-dropdown/style.css';
import pnp, { Web } from "sp-pnp-js";
import { INyRpgAdminSearchReactState } from './INyRpgAdminSearchReactState';
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/site-groups";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import autoBind from 'react-autobind';
import ReactHTMLTableToExcel from 'react-html-table-to-excel';
import { Pagination } from "@pnp/spfx-controls-react/lib/pagination";
//import 'office-ui-fabric-react';
//import { initializeIcons } from 'office-ui-fabric-react';
import { initializeIcons } from '@uifabric/icons';
initializeIcons();
//import Pagination from 'office-ui-fabric-react-pagination';
import { Grid } from  'react-loader-spinner';
import BootstrapTable from 'react-bootstrap-table-next';
import paginationFactory from 'react-bootstrap-table2-paginator';

const columns_table = [{
  dataField: 'ClearedOn',
  text: 'Cleared On',
  sort: true
}, {
  dataField: 'Project',
  text: 'Project',
  sort: true
}, {
  dataField: 'Employee_Name.Title',
  text: 'Employee Name',
  sort: true
},{
  dataField: 'Employee_Name.EMail',
  text: 'Employee Email ID',
  sort: true
},{
  dataField: 'Project_Role',
  text: 'Project Role',
  sort: true
},{
  dataField: 'EmployeeTitle',
  text: 'Title',
  sort: true
},{
  dataField: 'Manager.Title',
  text: 'Manager',
  sort: true
},{
  dataField: 'Status',
  text: 'Status',
  sort: true
}
];

export default class NyRpg extends React.Component<INyRpgAdminSearchProps, INyRpgAdminSearchReactState> {

  constructor(props: INyRpgAdminSearchProps, state: INyRpgAdminSearchReactState) {
    super(props);
    autoBind(this);
    this.state = {
      projectVals: [],
      projectSelected: '',
      empEmail: '',
      userID: 0,
      tableinfo: [],
      pagesArray:[],
      totalPages: 0,
      tenItemsTable: [],
      display: true,
      projectLen: 0,
      hideWhole: false,
      isAdmin: true,
      loading: false,
      SortOrder: 'ascn',
      tableName: '',
    };
  }

  @autoBind
  public async componentDidMount(){
    /*await this._getBulkData();
    if(this.state.tableinfo.length > 0){
      this.setState({
        display: false,
      });
      this.validationMessage();
    }
    else{
      this.setState({
        display: true,
      });
      this.validationMessage();
    }*/
    //this.getCurrentUser();
    this.checkAdmins();
    const val = this._getProjectValues();
    
    this.setState({
      projectVals: val,
    });

  }

   @autoBind
   private async checkAdmins(){
    const admins = [];
    const groupID = this.props.AdminGroupID;
    const users = await sp.web.siteGroups.getById(groupID).users();
    for (let entry of users) {
      admins.push(entry['Id']);
    }
    console.log(admins);
    await sp.web.currentUser.get().then((response) => {
      var user = response["Id"];

      var j = 0;
      for(j=0; j<admins.length; ++j) {
        if(admins[j] === user) {
          console.log("User is an Admin");
          this.setState({
            isAdmin: false,
          });  
        }
    }
   });
  }
  


  @autoBind

  private async _getPeoplePickerItems(items: any[]) {
    if (items[0] === undefined){
      this.setState({
        userID: 0,
      });
    }
    else{
      var a = items[0].secondaryText;
      var result = await sp.web.ensureUser(a);

      this.setState({
        empEmail: a,
        userID: result.data.Id
      });

      console.log(this.state.userID);
    }
  }

  
  public _onSelect = (itemSelected: { value: any; }) =>{
    this.setState({
      projectSelected: itemSelected.value,
    });
  }

  private tableName(){
    var x = new Date()
    var z = x.getTime();
    var str = 'ProjectClearance_' + z;
    this.setState({
      tableName: str,
    });
  }

  private _getProjectValues(){
    const listVals = ['Select a Project'];
    let web = new Web(this.props.context.pageContext.web.absoluteUrl);
      web.lists.getByTitle("AppleRPGProjects").items.select("Title").get().then(
       (response: any[])=>{
        this.setState({
          projectLen: response.length,
        });
        if(this.state.projectLen === 0){
          this.setState({
            hideWhole: true,
          });
        }
        else{
          this.setState({
            hideWhole: false,
          });

          for (let entry of response) {
            listVals.push(entry['Title']); // 1, "string", false
          }
        }
        
       }
     );
    return listVals;
  }

  private async _getBulkData(){
  
      this.setState({
        loading: true
      });
    let web = new Web(this.props.context.pageContext.web.absoluteUrl);
     await web.lists.getByTitle("ProjectClearenceRequest").items.select("ClearedOn, ExhibitA, Project, Employee_Name/Title, Employee_Name/EMail, Manager/Title, Project_Role, EmployeeTitle, Status").expand("Employee_Name, Manager").getAll().then(
        (response: any[])=>{

          for (let i = 0; i < response.length; i++) {
            if(response[i]['ClearedOn'] === null){
              response[i]['ClearedOn'] = '';
            }
            else{
              var arr = response[i]['ClearedOn'].split('T')[0].split('-');
              var str = arr[1] + '-' + arr[2] + '-' + arr[0];
              response[i]['ClearedOn'] = str;
            }
            
          }
         
         this.setState({
           tableinfo: response,
           tenItemsTable: response.slice(0,10),
         });
 
         console.log(this.state.tableinfo);
       }
      );
      const pageSize = 10;
      const pageCount = Math.ceil(this.state.tableinfo.length/pageSize);
      this.setState({
        totalPages: pageCount,
        loading: false,
      });
      
  }


  private async _getfilteredlistValues(){

    if (this.state.userID == 0 && (this.state.projectSelected === '' || this.state.projectSelected === 'Select a Project')){
      this.setState({
        loading: true
      });
    let web = new Web(this.props.context.pageContext.web.absoluteUrl);
     await web.lists.getByTitle("ProjectClearenceRequest").items.select("ClearedOn, ExhibitA, Project, Employee_Name/Title, Employee_Name/EMail, Manager/Title, Project_Role, EmployeeTitle, Status").expand("Employee_Name, Manager").getAll().then(
        (response: any[])=>{

          for (let i = 0; i < response.length; i++) {
            if(response[i]['ClearedOn'] === null){
              response[i]['ClearedOn'] = '';
            }
            else{
              var arr = response[i]['ClearedOn'].split('T')[0].split('-');
              var str = arr[1] + '-' + arr[2] + '-' + arr[0];
              response[i]['ClearedOn'] = str;
            }
            
          }
         
         this.setState({
           tableinfo: response,
           tenItemsTable: response.slice(0,10),
         });
 
         console.log(this.state.tableinfo);
       }
      );
      const pageSize = 10;
      const pageCount = Math.ceil(this.state.tableinfo.length/pageSize);
      this.setState({
        totalPages: pageCount,
        loading: false,
      });
    }
    
    else if (this.state.userID == 0){
      this.setState({
        loading: true
      });
    let web = new Web(this.props.context.pageContext.web.absoluteUrl);
     await web.lists.getByTitle("ProjectClearenceRequest").items.select("ClearedOn, ExhibitA, Project, Employee_Name/Title, Employee_Name/EMail, Manager/Title, Project_Role, EmployeeTitle, Status").expand("Employee_Name, Manager").filter("Project eq '" + this.state.projectSelected + "' ").getAll().then(
        (response: any[])=>{

          for (let i = 0; i < response.length; i++) {
            if(response[i]['ClearedOn'] === null){
              response[i]['ClearedOn'] = '';
            }
            else{
              var arr = response[i]['ClearedOn'].split('T')[0].split('-');
              var str = arr[1] + '-' + arr[2] + '-' + arr[0];
              response[i]['ClearedOn'] = str;
            }
            
          }
         
         this.setState({
           tableinfo: response,
           tenItemsTable: response.slice(0,10),
         });
 
         console.log(this.state.tableinfo);
       }
      );
      const pageSize = 10;
      const pageCount = Math.ceil(this.state.tableinfo.length/pageSize);
      this.setState({
        totalPages: pageCount,
        loading: false,
      });
      
    }

    else if (this.state.projectSelected === '' || this.state.projectSelected === 'Select a Project' ){
      this.setState({
        loading: true
      });
      let web = new Web(this.props.context.pageContext.web.absoluteUrl);

     await web.lists.getByTitle("ProjectClearenceRequest").items.select("ClearedOn, ExhibitA, Project, Employee_Name/Title, Employee_Name/EMail, Manager/Title, Project_Role, EmployeeTitle, Status").expand("Employee_Name, Manager").filter("Employee_NameId eq '" + this.state.userID + "' ").getAll().then(
        (response: any[])=>{

          for (let i = 0; i < response.length; i++) {
            if(response[i]['ClearedOn'] === null){
              response[i]['ClearedOn'] = '';
            }
            else{
              var arr = response[i]['ClearedOn'].split('T')[0].split('-');
              var str = arr[1] + '-' + arr[2] + '-' + arr[0];
              response[i]['ClearedOn'] = str;
            }
            
          }
         
         this.setState({
           tableinfo: response,
           tenItemsTable: response.slice(0,10),
         });
 
         console.log(this.state.tableinfo);
       }
      );
      const pageSize = 10;
      const pageCount = Math.ceil(this.state.tableinfo.length/pageSize);
      this.setState({
        totalPages: pageCount,
        loading: false,
      });
      
    }

    else{
      this.setState({
        loading: true
      });
      let web = new Web(this.props.context.pageContext.web.absoluteUrl);
     await web.lists.getByTitle("ProjectClearenceRequest").items.select("ClearedOn, ExhibitA, Project, Employee_Name/Title, Employee_Name/EMail, Manager/Title, Project_Role, EmployeeTitle, Status").expand("Employee_Name, Manager").filter("Project eq '" + this.state.projectSelected + "'  and Employee_NameId eq '" + this.state.userID + "' ").getAll().then(
       (response: any[])=>{

        for (let i = 0; i < response.length; i++) {
          if(response[i]['ClearedOn'] === null){
            response[i]['ClearedOn'] = '';
          }
          else{
            var arr = response[i]['ClearedOn'].split('T')[0].split('-');
            var str = arr[1] + '-' + arr[2] + '-' + arr[0];
            response[i]['ClearedOn'] = str;
          }
          
        }
        
        this.setState({
          tableinfo: response,
          tenItemsTable: response.slice(0,10),
        });
      }
     );
     const pageSize = 10;
     const pageCount = Math.ceil(this.state.tableinfo.length/pageSize);
     this.setState({
      totalPages: pageCount,
      loading: false,
    });

    }

    if(this.state.tableinfo.length > 0){
      this.setState({
        display: false,
      });
      this.validationMessage();
    }
    else{
      this.setState({
        display: true,
      });
      this.validationMessage();
    }
  
    this.tableName()
  }

  private _getPage(page: number){
    //console.log('Page:', page);
    var end = page*10;
    var start = end-10;
    //console.log(this.state.tableinfo.slice(start,end));
    this.setState({
      tenItemsTable: this.state.tableinfo.slice(start,end),
    });
  }

  private validationMessage(){
    if(this.state.tableinfo.length > 1){
    var markup = "<p> " + this.state.tableinfo.length + " entries found</p>";
    $("#greenValidation").html(markup);
    $("#redValidation").empty();
    }
    else if(this.state.tableinfo.length === 1){
    var markup1 = "<p> One entry found</p>";
    $("#greenValidation").html(markup1);
    $("#redValidation").empty();
    }
    else{
    var markup3 = "<p> No entries found</p>";
    $("#redValidation").html(markup3);
    $("#greenValidation").empty();
    }
  }


  public render(): React.ReactElement<INyRpgAdminSearchProps> {

    return (
      <>
      <div className='out-border'>
        <br></br>
      <div className="container container_width" hidden={this.state.hideWhole}>
        <header> NY-RPG Project Clearance Requests </header>
        <br />
        <div className="row">
        <div className="col-sm-1 tasksInput">
            <label htmlFor="employee" className='bold_label'>Employee:</label>
        </div>
        <div className="col-sm-5">
        <PeoplePicker
            context={this.props.context as any}
            titleText=""
            personSelectionLimit={1}
            groupName={""} // Leave this blank in case you want to filter from all users
            required={false}
            disabled={false}
            onChange={this._getPeoplePickerItems}
            showHiddenInUI={false}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000} />
          </div>

          <div className="col-sm-1 tasksInput">
            <label htmlFor="project" className='bold_label'>Project:</label>
        </div>
        <div className="col-sm-3">
        <Dropdown options={this.state.projectVals} onChange={this._onSelect} placeholder="Select a Project" />
        </div>
        <div className="col-sm-2">
        <button type='button' onClick={this._getfilteredlistValues} className="p_left btn btn-primary"> Search </button>
        </div>
        </div>
        <br></br>
        <div className="row">
          <div className="col-sm-5">
          
          </div>
          <div className="col-sm-2">
            {
            this.state.loading &&
            <Grid height="75" width="75" color='grey' ariaLabel='loading'/>
            }
          </div>
          <div className="col-sm-5">
          
          </div>
         
        </div>
        <div className="row">
          <div className="col-sm-6">
            <div id="greenValidation" className="greenvalidation">

            </div>
            <div id="redValidation" className="redvalidation">

            </div>
          
          </div>
          <div className="col-sm-6" hidden={(this.state.isAdmin || this.state.display)}>
            <div className="p_left">
            <img src="https://upload.wikimedia.org/wikipedia/commons/thumb/3/34/Microsoft_Office_Excel_%282019%E2%80%93present%29.svg/826px-Microsoft_Office_Excel_%282019%E2%80%93present%29.svg.png" alt="xls" width="30" height="30"></img>
            &nbsp;
            &nbsp;
          <ReactHTMLTableToExcel
            id="test-table-xls-button"
            className="download-table-xls-button btn btn-success"
            table="table-to-xls"
            filename={this.state.tableName}
            sheet="Project Clearance"
            buttonText="Download as XLS"
            />
            </div>
          </div>
          {/*<table className="table" 
          onCopy={(e)=>{
            e.preventDefault()
            return false;
          }}
          onCut={(e)=>{
            e.preventDefault()
            return false;
          }}
          onSelect={(e)=>{
            e.preventDefault()
            return false;
          }}
          hidden={this.state.display}>
              <thead>
                <br></br>
                <tr>
                  <th>Cleared On</th>
                  <th>Exhibit A</th>
                  <th>Project</th>
                  <th>Employee Name</th>
                  <th>Employee Email ID</th>
                  <th>Project Role</th>
                  <th>Title</th>
                  <th>Manager</th>
                  <th>Status</th>
                </tr>
              </thead>
              <tbody hidden={this.state.display}>
                {
                  this.state.tenItemsTable.map((res)=>
                  
                  <tr>
                    <td>{res.ClearedOn}</td>
                    <td>{res.ExhibitA}</td>
                    <td>{res.Project}</td>
                    <td>{res.Employee_Name.Title}</td>
                    <td>{res.Employee_Name.EMail}</td>
                    <td>{res.Project_Role}</td>
                    <td>{res.EmployeeTitle}</td>
                    <td>{res.Manager.Title}</td>
                    <td>{res.Status}</td>
                  </tr>
                  )
                }
              </tbody>
            </table>

              <div hidden={this.state.display}>
              <Pagination
                currentPage={1}
                totalPages={this.state.totalPages} 
                onChange={(page) => this._getPage(page)}
              />

              </div>
            

              </div>*/}
        <div className="row" hidden={true}>

            <table id="table-to-xls" className="table">
              <thead>
                <br></br>
                <tr>
                  <th>Cleared On</th>
                  <th>Exhibit A</th>
                  <th>Project</th>
                  <th>Employee Name</th>
                  <th>Employee Email ID</th>
                  <th>Project Role</th>
                  <th>Title</th>
                  <th>Manager</th>
                  <th>Status</th>
                </tr>
              </thead>
              <tbody>
                {
                  this.state.tableinfo.map((res)=>
                  
                  <tr>
                    <td>{res.ClearedOn}</td>
                    <td>{res.ExhibitA}</td>
                    <td>{res.Project}</td>
                    <td>{res.Employee_Name.Title}</td>
                    <td>{res.Employee_Name.EMail}</td>
                    <td>{res.Project_Role}</td>
                    <td>{res.EmployeeTitle}</td>
                    <td>{res.Manager.Title}</td>
                    <td>{res.Status}</td>
                  </tr>
                  )
                }
              </tbody>
            </table>
              </div>
        <br />
        <br></br>
        <div hidden={this.state.display}
            onCopy={(e)=>{
              e.preventDefault()
              return false;
            }}
            onCut={(e)=>{
              e.preventDefault()
              return false;
            }}
            onSelect={(e)=>{
              e.preventDefault()
              return false;
            }}>
          <BootstrapTable 
          keyField='id' 
          data={ this.state.tableinfo} 
          columns={ columns_table } 
          pagination={ paginationFactory() }
          />
        </div>
      </div>
      </div>
      </div>
          </>
    );
  }
}
