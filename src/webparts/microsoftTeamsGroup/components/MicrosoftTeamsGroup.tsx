import * as React from 'react';
import styles from './MicrosoftTeamsGroup.module.scss';
import { IMicrosoftTeamsGroupProps } from './IMicrosoftTeamsGroupProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp';
import {
  constructKeytip,
  DatePicker,
  DefaultButton,
  DetailsList,
  Dialog,
  Dropdown,
  IColumn,
  Icon,
  IIconProps,
  Label,
  PrimaryButton,
  SearchBox,
  TextField,
  ThemeSettingName,
  TooltipHost
} from 'office-ui-fabric-react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
// import { SelectTeamPicker } from "@pnp/spfx-controls-react/lib/TeamPicker";



export interface IMicrosoftTeamsGroupState {
  TaskDetailsData: any;
  AddTaskDialogOpen: boolean;
  AllTaskListDetails: any;
  TaskName: any;
  Description: any;
  StartDate: any;
  EndDate: any;
  ProjectManagerID: any;
  ProjectManager: any;
  Status: any;
  AssignedTo: any;
  ProjectNameID: any;
  ProjectName: any;
  AssignedToID: any;
  Statuslist: any;
  TaskFormSection1: boolean;
  TaskFormSection2: boolean;
  TaskFormSection3: boolean;
  SelectedProjectNamelist: any;
  ProjectManagerlist: any;
  ProjectNamelist: any;
  EditTaskName: any;
  EditDescription: any;
  EditStartDate: any;
  EditEndDate: any;
  EditProjectManagerID: any;
  EditProjectManager: any;
  EditStatus: any;
  EditAssignedTo: any;
  EditProjectNameID: any;
  EditProjectName: any;
  EditAssignedToID: any;
  EditTaskDialogOpen: boolean;
  DeleteTaskDialogOpen: boolean;
  CurrentTaskDetailsID : any;
  DeleteTaskDetailsID : any;
}


const addIcon: IIconProps = { iconName: 'Add' };

const SendIcon: IIconProps = { iconName: 'Send' };

const CancelIcon: IIconProps = { iconName: 'Cancel' };

const DeleteIcon: IIconProps = { iconName: 'Delete' };

const TextDocumentEdit: IIconProps = { iconName: 'TextDocumentEdit' };

const AddTaskDetailsDialogContentProps = {
  title: "Add Task Details",
};

const ReadTaskDetailsDialogContentProps = {
  title: "Read Task Details"
};

const UpdateTaskDetailsDialogContentProps = {
  title: "Update Task Details"
};

const DeleteTaskDetailsFilterDialogContentProps = {
  title: "Confirm Deletion"
};


const addmodelProps = {
  className: "Add-Dialog"
};

const readmodelProps = {
  className: "Read-Dialog"
};

const updatemodelProps = {
  className: "Update-Dialog"
};

const deletmodelProps = {
  className: "Delet-Dialog"
};

require("../assets/css/style.css");
require("../assets/css/fabric.min.css");

export default class MicrosoftTeamsGroup extends React.Component<IMicrosoftTeamsGroupProps, IMicrosoftTeamsGroupState> {

  constructor(props: IMicrosoftTeamsGroupProps, state: IMicrosoftTeamsGroupState) {

    super(props);

    this.state = {
      TaskDetailsData: [],
      AddTaskDialogOpen: true,
      AllTaskListDetails: [],
      TaskName: "",
      Description: "",
      StartDate: "",
      EndDate: "",
      ProjectManagerID: "",
      ProjectManager: "",
      Status: "",
      AssignedTo: "",
      AssignedToID: "",
      ProjectNameID: "",
      ProjectName: "",
      Statuslist: [],
      TaskFormSection1: true,
      TaskFormSection2: false,
      TaskFormSection3: false,
      SelectedProjectNamelist: [],
      ProjectManagerlist: [],
      ProjectNamelist: [],
      EditTaskName: "",
      EditDescription: "",
      EditStartDate: "",
      EditEndDate: "",
      EditProjectManagerID: "",
      EditProjectManager: "",
      EditStatus: "",
      EditAssignedTo: "",
      EditProjectNameID: "",
      EditProjectName: "",
      EditAssignedToID: "",
      EditTaskDialogOpen: true,
      DeleteTaskDialogOpen: true,
      CurrentTaskDetailsID : "",
      DeleteTaskDetailsID : ""
    };

  }

  public render(): React.ReactElement<IMicrosoftTeamsGroupProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    const columns: IColumn[] = [
      {
        key: "TaskName",
        name: "Task Name",
        fieldName: "TaskName",
        minWidth: 150,
        maxWidth: 150,
        isResizable: false
      },
      {
        key: "StartDate",
        name: "Start Date",
        fieldName: "StartDate",
        minWidth: 130,
        maxWidth: 130,
        isResizable: false
      },
      {
        key: "ProjectManager",
        name: "Project Manager",
        fieldName: "ProjectManager",
        minWidth: 150,
        maxWidth: 150,
        isResizable: false
      },
      {
        key: "ProjectID",
        name: "ProjectName",
        fieldName: "ProjectID",
        minWidth: 150,
        maxWidth: 150,
        isResizable: false
      },
      {
        key: "Status",
        name: "Task Status",
        fieldName: "Status",
        minWidth: 150,
        maxWidth: 150,
        isResizable: false
      },
      {
        key: "AssignedTo",
        name: "AssignedTo",
        fieldName: "AssignedTo",
        minWidth: 150,
        maxWidth: 150,
        isResizable: false
      },
      {
        key: "Actions",
        name: "Actions",
        fieldName: "Actions",
        minWidth: 150,
        maxWidth: 150,
        isResizable: false,
        onRender: (item) => {
          return (
            <div>
              <div className='ms-Grid-row'>
                <div className='ms-Grid-col'>
                  <div className='TaskAction-Icon'>

                    <div className='Read-Icon'>
                      <Icon iconName='View' className='Read-task' ></Icon>
                    </div>

                    <div className='Edit-Icon'>
                      <Icon className='Edit-Icon' iconName="Edit" onClick={() => this.setState({EditTaskDialogOpen : false , CurrentTaskDetailsID : item.ID }, () => this.GetEditTaskDetails(item.ID))}></Icon>
                    </div>

                    <div className='Delete-Icon'>
                      <Icon className='icon' iconName="Delete" onClick={() => this.setState({ DeleteTaskDialogOpen : false , DeleteTaskDetailsID: item.ID })}></Icon>
                    </div>

                  </div>
                </div>
              </div>
            </div>

          )
        }
      }
    ];


    return (

      <section id="microsoftTeamsGroup">
        <div className='ms-Grid'>

          <div className='Task-Header'>
            <h3>Task Details</h3>
          </div>

          <div className='ms-Grid-row'>
            <div className='Task-Group'>

              <div className='ms-Grid-col ms-sm5 ms-md4 ms-lg2'>
                <SearchBox placeholder="Search" className="new-search"
                  onChange={(e) => { this.applyVendorFilters(e.target.value); }}
                  onClear={(e) => { this.applyVendorFilters(e.target.value); }}
                />
              </div>

              <div className='ms-Grid-col ms-sm1 ms-md1 ms-lg10 Add-Tasks'>
                <div className='Add-Details'>
                  <PrimaryButton
                    iconProps={addIcon}
                    text="Add Task"
                    onClick={() => this.setState({ AddTaskDialogOpen: false })}
                  />
                </div>
              </div>

            </div>
          </div>

          <Dialog
            hidden={this.state.AddTaskDialogOpen}
            onDismiss={() =>
              this.setState({
                TaskName: "",
                Description: "",
                StartDate: "",
                EndDate: "",
                ProjectManagerID: "",
                Status: "",
                AssignedTo: "",
                ProjectNameID: "",
                AddTaskDialogOpen: true,
                TaskFormSection1: true,
                TaskFormSection2: false,
                TaskFormSection3: false
              })
            }
            dialogContentProps={AddTaskDetailsDialogContentProps}
            modalProps={addmodelProps}
            minWidth={500}
          >

            <div className='ms-Grid-row'>

              <div>
                {
                  this.state.TaskFormSection1 == true ?
                    <>
                      <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                        <div className='Add-TaskName'>
                          <TextField
                            label="Task Name"
                            type="text"
                            required={true}
                            onChange={(value) =>
                              this.setState({ TaskName: value.target["value"] })
                            }
                          />
                        </div>
                      </div>

                      <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                        <div className='Add-Description'>
                          <TextField
                            label='Description'
                            type='text'
                            multiline rows={3}
                            required={true}
                            onChange={(value) =>
                              this.setState({ Description: value.target["value"] })
                            }
                          />
                        </div>
                      </div>

                      <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                        <div className='Add-StartDate'>
                          <DatePicker
                            label='Start Date'
                            allowTextInput={false}
                            value={this.state.StartDate ? this.state.StartDate : null}
                            onSelectDate={(date: any) => this.setState({ StartDate: date })}
                            ariaLabel="Select a Start date" placeholder="Select a Start date" isRequired />
                        </div>
                      </div>

                      <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                        <div className='Add-EndDate'>
                          <DatePicker
                            label='End Date'
                            allowTextInput={false}
                            value={this.state.EndDate ? this.state.EndDate : null}
                            onSelectDate={(date: any) => this.setState({ EndDate: date })}
                            ariaLabel="Select a End date" placeholder="Select a End date" isRequired />
                        </div>
                      </div>

                      <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                        <div className='Next'>
                          <PrimaryButton
                            text="Next"
                            onClick={() => this.setState({ TaskFormSection1: false, TaskFormSection2: true })}
                          />
                        </div>
                      </div>

                    </>
                    :
                    <>
                      <div>
                        {
                          this.state.TaskFormSection2 == true ?
                            <>
                              <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                                <div className='add-ProjectManager'>
                                  <Dropdown
                                    options={this.state.ProjectManagerlist}
                                    placeholder="Select Your Project acording Project Maneger"
                                    label="Project Manager"
                                    required
                                    onChange={(e, option, text) =>
                                      this.setState({ ProjectManager: option.text, ProjectManagerID: option.key })
                                    }
                                  />
                                </div>
                              </div>

                              <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                                <div className='add-ProjectName'>
                                  <Dropdown
                                    options={this.state.ProjectNamelist}
                                    placeholder="Select Your Project"
                                    label="Project Name"
                                    required
                                    onChange={(e, option, text) =>
                                      this.setState({ ProjectName: option.text, ProjectNameID: option.key })
                                    }
                                  />
                                </div>
                              </div>

                              <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                                <div className='add-Status'>
                                  <Dropdown
                                    options={this.state.Statuslist}
                                    placeholder="Select Your Task Status"
                                    label="Task Status"
                                    required
                                    onChange={(e, option, text) =>
                                      this.setState({ Status: option.text })
                                    }
                                  />
                                </div>
                              </div>

                              <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                                <div className='Next'>
                                  <PrimaryButton
                                    text="Next"
                                    onClick={() => this.setState({ TaskFormSection2: false, TaskFormSection3: true })}
                                  />
                                </div>
                              </div>

                            </>
                            :
                            <>
                              <div>
                                {
                                  this.state.TaskFormSection3 == true ?
                                    <>
                                      <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                                        <div className='add-AssignedTo'>
                                          <PeoplePicker
                                            context={this.props.context}
                                            titleText="Assigned To:"
                                            personSelectionLimit={1}
                                            placeholder='Select Assigned To'
                                            showtooltip={true}
                                            required={true}
                                            defaultSelectedUsers={[this.state.AssignedTo.Title]}
                                            onChange={(e) =>
                                              this.setState({ AssignedToID: e[0].id, AssignedTo: e[0].text })
                                            }
                                            principalTypes={[PrincipalType.User]}
                                            resolveDelay={300}
                                            ensureUser={true}
                                          />
                                        </div>
                                      </div>

                                      <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                                        <div className='Submit-TaskDetails'>
                                          <PrimaryButton
                                            className='Save-Details'
                                            text="Submit"
                                            onClick={() => this.AddTaskDetails()}
                                          />
                                        </div>
                                      </div>

                                    </>
                                    :
                                    <></>
                                }
                              </div>
                            </>
                        }
                      </div>
                    </>
                }
              </div>
            </div>

          </Dialog>

          <Dialog
            hidden={this.state.EditTaskDialogOpen}
            onDismiss={() =>
              this.setState({
                EditTaskName: "",
                EditDescription: "",
                EditStartDate: "",
                EditEndDate: "",
                EditProjectManagerID: "",
                EditStatus: "",
                EditAssignedTo: "",
                EditProjectNameID: "",
                EditTaskDialogOpen: true,
                TaskFormSection1: true,
                TaskFormSection2: false,
                TaskFormSection3: false
              })
            }
          >
            <div className='ms-Grid-row'>

              <div>
                {
                  this.state.TaskFormSection1 == true ?
                    <>
                      <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                        <div className='Edit-TaskName'>
                          <TextField
                            label="Task Name"
                            type="text"
                            required={true}
                            onChange={(value) =>
                              this.setState({ EditTaskName: value.target["value"] })
                            }
                          />
                        </div>
                      </div>

                      <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                        <div className='Edit-Description'>
                          <TextField
                            label='Description'
                            type='text'
                            multiline rows={3}
                            required={true}
                            onChange={(value) =>
                              this.setState({ EditDescription: value.target["value"] })
                            }
                          />
                        </div>
                      </div>

                      <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                        <div className='Edit-StartDate'>
                          <DatePicker
                            label='Start Date'
                            allowTextInput={false}
                            value={this.state.EditStartDate ? this.state.EditStartDate : null}
                            onSelectDate={(date: any) => this.setState({ EditStartDate: date })}
                            ariaLabel="Select a Start date" placeholder="Select a Start date" isRequired />
                        </div>
                      </div>

                      <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                        <div className='Add-EndDate'>
                          <DatePicker
                            label='End Date'
                            allowTextInput={false}
                            value={this.state.EditEndDate ? this.state.EditEndDate : null}
                            onSelectDate={(date: any) => this.setState({ EditEndDate: date })}
                            ariaLabel="Select a End date" placeholder="Select a End date" isRequired />
                        </div>
                      </div>

                      <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                        <div className='Next'>
                          <PrimaryButton
                            text="Next"
                            onClick={() => this.setState({ TaskFormSection1: false, TaskFormSection2: true })}
                          />
                        </div>
                      </div>
                    </>
                    :
                    <>
                      <div>
                        {
                          this.state.TaskFormSection2 == true ?
                            <>
                              <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                                <div className='edit-ProjectManager'>
                                  <Dropdown
                                    options={this.state.ProjectManagerlist}
                                    placeholder="Select Your Project acording Project Maneger"
                                    label="Project Manager"
                                    required
                                    defaultSelectedKey={this.state.EditProjectManagerID}
                                    onChange={(e, option, text) =>
                                      this.setState({ EditProjectManager: option.text, EditProjectManagerID: option.key })
                                    }
                                  />
                                </div>
                              </div>

                              <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                                <div className='add-ProjectName'>
                                  <Dropdown
                                    options={this.state.ProjectNamelist}
                                    placeholder="Select Your Project"
                                    label="Project Name"
                                    required
                                    defaultSelectedKey={this.state.EditProjectNameID}
                                    onChange={(e, option, text) =>
                                      this.setState({ EditProjectName: option.text, EditProjectNameID: option.key })
                                    }
                                  />
                                </div>
                              </div>

                              <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                                <div className='add-Status'>
                                  <Dropdown
                                    options={this.state.Statuslist}
                                    placeholder="Select Your Task Status"
                                    label="Task Status"
                                    required
                                    defaultSelectedKey={this.state.EditStatus}
                                    onChange={(e, option, text) =>
                                      this.setState({ EditStatus: option.text })
                                    }
                                  />
                                </div>
                              </div>

                              <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                                <div className='Next'>
                                  <PrimaryButton
                                    text="Next"
                                    onClick={() => this.setState({ TaskFormSection2: false, TaskFormSection3: true })}
                                  />
                                </div>
                              </div>
                            </>
                            :
                            <>
                              <div>
                                  {
                                    this.state.TaskFormSection3 == true ?
                                      <>
                                        <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                                          <div className='edit-AssignedTo'>
                                            <PeoplePicker
                                              context={this.props.context}
                                              titleText="Assigned To:"
                                              personSelectionLimit={1}
                                              placeholder='Select Assigned To'
                                              showtooltip={true}
                                              required={true}
                                              defaultSelectedUsers={[this.state.EditAssignedTo.Title]}
                                              onChange={(e) =>
                                                this.setState({ EditAssignedToID: e[0].id, EditAssignedTo: e[0].text })
                                              }
                                              principalTypes={[PrincipalType.User]}
                                              resolveDelay={300}
                                              ensureUser={true}
                                            />
                                          </div>
                                        </div>

                                        <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                                          <div className='Submit-TaskDetails'>
                                            <PrimaryButton
                                              className='Save-Details'
                                              text="Update Task Details"
                                              onClick={() => this.UpateTaskDetails(this.state.CurrentTaskDetailsID)}
                                            />
                                          </div>
                                        </div>

                                      </>
                                      :
                                      <></>
                                  }
                              </div>
                            </>
                        }
                      </div>
                    </>
                }
              </div>

            </div>

          </Dialog>

          <Dialog
          hidden={this.state.DeleteTaskDialogOpen}
          onDismiss={() =>
            this.setState({
              DeleteTaskDialogOpen: true
            })
          }
          dialogContentProps={DeleteTaskDetailsFilterDialogContentProps}
          modalProps={deletmodelProps}
          minWidth={500}
        >

          <div className="DeleteClose-Icon">
            <div className='delete-text'>
              {/* <h5 className='confirm-text'>Confirm Deletion</h5> */}
              <Icon iconName="Cancel" className='confirm-icon' onClick={() => this.setState({ DeleteTaskDialogOpen: true })}></Icon>
            </div>
            <div className="delete-msg">
              <Icon iconName='Warning' className='Warinig-Ic'></Icon>
              <p className='mb-0'>Are you sure? <br />Do you really want to delete these record? </p>
            </div>
            <div className='Delet-buttons'>
              <DefaultButton
                className="cancel-Icon"
                text='Cancel'
                iconProps={CancelIcon}
                onClick={() => this.setState({ DeleteTaskDialogOpen: true })}
              />

              <PrimaryButton
                className='delete-icon'
                text='Delete'
                iconProps={DeleteIcon}
                onClick={() => this.DeleteTaskDetails(this.state.DeleteTaskDetailsID)}
              />
            </div>
          </div>

        </Dialog>

          <div className='ms-Grid'>
            <DetailsList
              className="TaskDetails-List"
              items={this.state.TaskDetailsData}
              columns={columns}
              setKey='set'
              layoutMode={1}
              selectionMode={0}
              isHeaderVisible={true}
              ariaLabelForSelectionColumn="Toggle selection"
              ariaLabelForSelectAllCheckbox="Toggle selection for all items"
              checkButtonAriaLabel="select row"
            >
            </DetailsList>
          </div>


        </div>
      </section>

    );
  }

  public async componenettDidMount() {
    this.GetTaskDetailsItems();
    this.GetTaskdetailsStatusItems();
    // this._getPeoplePickerItems;
  }

  public async GetTaskDetailsItems() {
    const tasks = await sp.web.lists.getByTitle("Project Task list").items.select(
      "ID",
      "TaskName",
      "Description",
      "StartDate",
      "EndDate",
      "ProjectManager/ID",
      "ProjectManager/ProjectManager",
      "Status",
      "AssignedTo/Title",
      "AssignedTo/Id",
      "ProjectName/ID",
      "ProjectName/ProjectName"
    ).expand("AssignedTo", "ProjectName", "ProjectManager").get().then((data) => {
      let AllData = [];
      console.log(data);
      console.log(tasks);

      if (data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID: item.Id ? item.Id : "",
            TaskName: item.TaskName ? item.TaskName : "",
            Description: item.Description ? item.Description : "",
            StartDate: item.StartDate ? item.StartDate : "",
            EndDate: item.EndDate ? item.EndDate : "",
            ProjectManagerId: item.ProjectManager ? item.ProjectManager.ID : "",
            ProjectManager: item.ProjectManager ? item.ProjectManager.ProjectManager : "",
            Status: item.Status ? item.Status : "",
            AssignedTo: item.AssignedTo ? item.AssignedTo.Title : "",
            ProjectNameId: item.ProjectName ? item.ProjectName.ID : "",
            ProjectName: item.ProjectName ? item.ProjectName.ProjectName : ""
          });
        });
        this.setState({ TaskDetailsData: AllData, AllTaskListDetails: AllData });
        console.log(this.state.TaskDetailsData);
      }
    }).catch((error) => {
      console.log("Error while fetching task details: ", error);
    });
  }

  private async applyVendorFilters(Test) {
    if (Test) {
      let SerchText = Test.toLowerCase();

      let filteredData = this.state.AllTaskListDetails.filter((x) => {
        let CompanyName = x.CompanyName.toLowerCase();
        let CompanyEmail = x.CompanyEmail.toLowerCase();
        return (
          CompanyName.includes(SerchText) || CompanyEmail.includes(SerchText)
        );
      });

      this.setState({ TaskDetailsData: filteredData });
    }
    else {
      this.setState({ TaskDetailsData: this.state.AllTaskListDetails });
    }
  }

  public async AddTaskDetails() {
    if (this.state.TaskName.length = 0) {
      alert("Please enter task name");
    }
    else {
      const addtaskdetails = await sp.web.lists.getByTitle("Project Task list").items.add({
        TaskName: this.state.TaskName,
        Description: this.state.Description,
        StartDate: this.state.StartDate,
        EndDate: this.state.EndDate,
        Status: this.state.Status,
        ProjectNameId: this.state.ProjectNameID,
        ProjectManagerId: this.state.ProjectManagerID,
        AssignedToId: this.state.AssignedToID
      }).catch((error) => {
        console.log("Can't add a new task", error);
      });

      this.GetTaskDetailsItems();
      this.setState({ TaskDetailsData: addtaskdetails });
      this.setState({ AddTaskDialogOpen: true });
      this.setState({ TaskFormSection1: true, TaskFormSection2: false, TaskFormSection3: false });
    }
  }

  public async GetEditTaskDetails(ID) {
    let Edittaskdetails = this.state.TaskDetailsData.filter((item) => {
      if (item.ID == ID) {
        return item;
      }
    });
    console.log(Edittaskdetails);
    this.setState({
      EditTaskName: Edittaskdetails[0].TaskName,
      EditDescription: Edittaskdetails[0].Description,
      EditStartDate: Edittaskdetails[0].StartDate,
      EditEndDate: Edittaskdetails[0].EndDate,
      EditProjectManagerID: Edittaskdetails[0].ProjectManagerId,
      // EditProjectManager: Edittaskdetails[0].ProjectManager,
      EditStatus: Edittaskdetails[0].Status,
      EditAssignedTo: Edittaskdetails[0].AssignedTo,
      EditProjectNameID: Edittaskdetails[0].ProjectNameId,
      // EditProjectName: Edittaskdetails[0].ProjectName,
      // EditAssignedToID: Edittaskdetails[0].AssignedToID
    });

  }

  public async UpateTaskDetails(CurrentTaskDetailsID) {
    const updatetaskdetails = await sp.web.lists.getByTitle("Project Task list").items.getById(CurrentTaskDetailsID).update({
      TaskName: this.state.EditTaskName,
      Description: this.state.EditDescription,
      StartDate: this.state.EditStartDate,
      EndDate: this.state.EditEndDate,
      Status: this.state.EditStatus,
      ProjectNameId: this.state.EditProjectNameID,
      ProjectManagerId: this.state.EditProjectManagerID,
      AssignedToId: this.state.EditAssignedToID
    }).catch((error) => {
      console.log(error);
    });
    this.GetTaskDetailsItems();
    this.setState({ EditTaskDialogOpen: true });
    this.setState({ TaskDetailsData: updatetaskdetails });
    this.setState({ TaskFormSection1: true, TaskFormSection2: false, TaskFormSection3: false })
  }

  public async DeleteTaskDetails(DeleteTaskDetailsID) {
    const deletetaskdetails = await sp.web.lists.getByTitle("Project Task list").items.getById(DeleteTaskDetailsID).delete();
    this.setState({ TaskDetailsData: deletetaskdetails });
    this.setState({ DeleteTaskDialogOpen: true });
    this.GetTaskDetailsItems();
  }

  public async GetTaskdetailsStatusItems() {
    const choieFieldName1 = "Status";
    const field1 = await sp.web.lists.getByTitle("Project Task list").fields.getByInternalNameOrTitle(choieFieldName1)();
    let status = [];
    field1["Choices"].forEach(function (dname, i) {
      status.push({ key: dname, text: i });
    });
    this.setState({ Statuslist: status });

    const choiceFieldName2 = "ProjectManager";
    const field2 = await sp.web.lists.getByTitle("Project Task list").fields.getByInternalNameOrTitle(choiceFieldName2)();
    let projectManager = [];
    field2["Choices"].forEach(function (dname, i) {
      projectManager.push({ key: dname, text: i });
    });
    this.setState({ ProjectManagerlist: projectManager });

    const choiceFieldName3 = "ProjectName";
    const field3 = await sp.web.lists.getByTitle("Project Task list").fields.getByInternalNameOrTitle(choiceFieldName3)();
    let projectName = [];
    field3["Choices"].forEach(function (dname, i) {
      projectName.push({ key: dname, text: i });
    });
    this.setState({ ProjectNamelist: projectName });
  }

  public _getPeoplePickerItems = async (items: any[]) => {
    if (items.length > 0) {
      const assigneto = items.map(item => item.text);
      const assignetoID = items.map(item => item.id);
      this.setState({ AssignedTo: assigneto });
      this.setState({ AssignedToID: assignetoID });
    }
    else {
      this.setState({ AssignedTo: [] });
      this.setState({ AssignedToID: [] });
    }
  }

  // public async handleProjectName(SelectedProjectName) {
  //   let project = this.state.TaskDetailsData;
  //   const selectedProjectname = project.filter((item) => {
  //     if (item.ProjectID == SelectedProjectName) {
  //       return item;
  //     }
  //   })
  // }

}
