import * as React from 'react';
import styles from './MicrosoftTeamsGroup.module.scss';
import { IMicrosoftTeamsGroupProps } from './IMicrosoftTeamsGroupProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp';
import {
  AutoScroll,
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
import * as moment from 'moment';
import Slider from "react-slick";
import "slick-carousel/slick/slick.css";
import "slick-carousel/slick/slick-theme.css";
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
  Priority: any;
  Prioritylist: any;
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
  EditPriority: any;
  EditAssignedTo: any;
  EditProjectNameID: any;
  EditProjectName: any;
  EditAssignedToID: any;
  EditTaskDialogOpen: boolean;
  DeleteTaskDialogOpen: boolean;
  CurrentTaskDetailsID: any;
  DeleteTaskDetailsID: any;
  requestID: any;
  ProjectDetailsData: any;
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
      TaskDetailsData: "",
      AddTaskDialogOpen: true,
      AllTaskListDetails: [],
      TaskName: "",
      Description: "",
      StartDate: "",
      EndDate: "",
      ProjectManagerID: "",
      ProjectManager: "",
      Status: [],
      AssignedTo: [],
      AssignedToID: [],
      ProjectNameID: "",
      ProjectName: "",
      Statuslist: [],
      Priority: "",
      Prioritylist: [],
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
      EditPriority: "",
      EditAssignedTo: [],
      EditProjectNameID: "",
      EditProjectName: "",
      EditAssignedToID: [],
      EditTaskDialogOpen: true,
      DeleteTaskDialogOpen: true,
      CurrentTaskDetailsID: "",
      DeleteTaskDetailsID: "",
      requestID: "",
      ProjectDetailsData: ""
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


    var settings = {
      dots: true,
      infinite: true,
      speed: 500,
      slidesToShow: 1,
      slidesToScroll: 1,
      autoplaySpeed: 2000,
      autoplay: true,
      cssEase: "linear",
      // nextArrow: <SampleNextArrow />,
      // prevArrow: <SamplePrevArrow />
    };

    function SampleNextArrow(props) {
      const { className, style, onClick } = props;
      return (
        <img className={className + " arrow-img-icon"} src={require("../assets/Image/next.jpg")} style={{ ...style }} onClick={onClick} />
      );
    }

    function SamplePrevArrow(props) {
      const { className, style, onClick } = props;
      return (
        <img className={className + " arrow-img-icon"} src={require("../assets/Image/back.jpg")} style={{ ...style }} onClick={onClick} />
      );
    }

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
        isResizable: false,
        onRender: (item) => {
          return <span>{moment(new Date(item.StartDate)).format("DD-MM-YYYY")}</span>;
        }
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
        key: "ProjectName",
        name: "Project Name",
        fieldName: "ProjectName",
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
        name: "Assigned To",
        fieldName: "AssignedTo",
        minWidth: 150,
        maxWidth: 150,
        isResizable: false,
        onRender: (item) => {
          return <span>
            {item.AssignedTo && item.AssignedTo.length > 0
              ? item.AssignedTo.map(member => member.Title).join(', ')
              : ''}
          </span>;
        }
        // {item.AssignedTo.Title || ''}
        //  onRender: (item) => {
        //   return (
        //     <span>
        //       {Array.isArray(item.AssignedTo) ? (
        //         item.AssignedTo.map((member, index) => (
        //           <div key={index}>
        //             <p>{member.Title}</p>
        //           </div>
        //         ))
        //       ) : item.AssignedTo ? (
        //         <div>
        //           <p>{item.AssignedTo.Title}</p>
        //         </div>
        //       ) : (
        //         ""
        //       )}
        //     </span>
        //   );
        // }
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
                      <Icon className='Edit-Icon' iconName="Edit" onClick={() => this.setState({ EditTaskDialogOpen: false, CurrentTaskDetailsID: item.ID }, () => this.GetEditTaskDetails(item.ID))}></Icon>
                    </div>

                    <div className='Delete-Icon'>
                      <Icon className='icon' iconName="Delete" onClick={() => this.setState({ DeleteTaskDialogOpen: false, DeleteTaskDetailsID: item.ID })}></Icon>
                    </div>

                  </div>
                </div>
              </div>
            </div>

          );
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
            <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
              <div className='Status-box'>
                {
                  this.state.TaskDetailsData.length > 0 &&
                  this.state.TaskDetailsData.map((item) => {
                    return (
                      item.Status == "To Do" ?
                        <>
                          <div className='Status-kit'>
                            <div className='Status'>
                              <h2 className='InProgress-Title'>To Do</h2>

                              <div className='Inprogress-card Status-wrapper'>
                                <h3 className="card-title">{item.TaskName}</h3>
                                <p>{item.Description}</p>
                                {
                                  item.Priority == "Low" ?
                                    <>
                                      <span className="badge-low">{item.Priority}</span>
                                    </>
                                    :
                                    <>
                                      {
                                        item.Priority == "Medium" ?
                                          <>
                                            <span className="badge-medium">{item.Priority}</span>
                                          </>
                                          :
                                          <>
                                            {
                                              item.Priority == "High" ?
                                                <>
                                                  <span className="badge-high">{item.Priority}</span>
                                                </>
                                                :
                                                <></>
                                            }
                                          </>
                                      }
                                    </>
                                }

                              </div>
                            </div>

                          </div>
                        </>
                        :
                        <>
                          {
                            item.Status == "In Progress" ?
                              <>
                                <div className='Status-Kit'>
                                  <div className='Status'>
                                    <h2 className='InProgress-Title'>In Progress</h2>

                                    <div className='Inprogress-card Status-wrapper'>
                                      <h3 className="card-title">{item.TaskName}</h3>
                                      <p>{item.Description}</p>
                                      {
                                        item.Priority == "Low" ?
                                          <>
                                            <span className="badge-low">{item.Priority}</span>
                                          </>
                                          :
                                          <>
                                            {
                                              item.Priority == "Medium" ?
                                                <>
                                                  <span className="badge-medium">{item.Priority}</span>
                                                </>
                                                :
                                                <>
                                                  {
                                                    item.Priority == "High" ?
                                                      <>
                                                        <span className="badge-high">{item.Priority}</span>
                                                      </>
                                                      :
                                                      <></>
                                                  }
                                                </>
                                            }
                                          </>
                                      }
                                    </div>
                                  </div>
                                </div>
                              </>
                              :
                              <>

                              </>
                          }
                        </>



                    );
                  })

                }
              </div>
            </div>
          </div>

          <br />

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
                Priority: "",
                AssignedTo: [],
                AssignedToID: [],
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
                              <div className='ms-Grid-col ms-sm6 ms-md6 ms-lg6'>
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

                              <div className='ms-Grid-col ms-sm6 ms-md6 ms-lg6'>
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

                              <div className='ms-Grid-col ms-sm6 ms-md6 ms-lg6'>
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

                              <div className='ms-Grid-col ms-sm6 ms-md6 ms-lg6'>
                                <div className="add-Priority">
                                  <Dropdown
                                    options={this.state.Prioritylist}
                                    placeholder='Select Your Task Priority'
                                    label='Task Priority'
                                    required
                                    onChange={(e, option, text) =>
                                      this.setState({ Priority: option.text })
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
                                            personSelectionLimit={3}
                                            placeholder='Select Assigned To'
                                            showtooltip={true}
                                            required={true}
                                            // defaultSelectedUsers={[this.state.AssignedTo.Title]}
                                            // onChange={(e) =>
                                            //    this.setState({ AssignedToID: e[0].id, AssignedTo: e[0].text })
                                            // }
                                            onChange={this._getPeoplePickerItems}
                                            principalTypes={[PrincipalType.User]}
                                            resolveDelay={300}
                                            ensureUser={true}
                                          />
                                        </div>
                                      </div>

                                      <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                                        <div className='Add-TaskDetails'>
                                          <div className='Submit-TaskDetails'>
                                            <PrimaryButton
                                              className='Save-Details'
                                              text="Submit"
                                              onClick={() => this.AddTaskDetails()}
                                            />
                                          </div>

                                          <div className='Cancel-Project'>
                                            <DefaultButton
                                              iconProps={CancelIcon}
                                              text="Cancel"
                                              onClick={() =>
                                                this.setState({ AddTaskDialogOpen: true, TaskFormSection1: true, TaskFormSection2: false, TaskFormSection3: false })
                                              }
                                            />
                                          </div>
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
                EditTaskDialogOpen: true,
                EditTaskName: "",
                EditDescription: "",
                EditStartDate: "",
                EditEndDate: "",
                EditProjectManagerID: "",
                EditStatus: [],
                EditPriority: [],
                EditAssignedTo: "",
                EditProjectNameID: "",
                TaskFormSection1: true,
                TaskFormSection2: false,
                TaskFormSection3: false
              })
            }
            dialogContentProps={UpdateTaskDetailsDialogContentProps}
            modalProps={updatemodelProps}
            minWidth={500}
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
                            value={this.state.EditTaskName}
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
                            value={this.state.EditDescription}
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
                              <div className='ms-Grid-col ms-sm6 ms-md6 ms-lg6'>
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

                              <div className='ms-Grid-col ms-sm6 ms-md6 ms-lg6'>
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

                              <div className='ms-Grid-col ms-sm6 ms-md6 ms-lg6'>
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

                              <div className='ms-Grid-col ms-sm6 ms-md6 ms-lg6'>
                                <div className='add-Priority'>
                                  <Dropdown
                                    options={this.state.Prioritylist}
                                    placeholder='Select Your Task Priority'
                                    label='Task Priority'
                                    required
                                    defaultSelectedKey={this.state.EditPriority}
                                    onChange={(e, option, text) =>
                                      this.setState({ EditPriority: option.text })
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
                                            defaultSelectedUsers={this.state.EditAssignedTo}
                                            // onChange={(e) =>
                                            //   this.setState({ EditAssignedToID: e[0].id, EditAssignedTo: e[0].text })
                                            // }
                                            onChange={this._getPeoplePickerItems}
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

  public async componentDidMount() {
    this.GetTaskdetailsStatusItems();
    this.GetTaskDetails();
  }

  public async GetTaskDetails() {
    try {
      const urlParams = new URLSearchParams(window.location.search);
      const requestid = urlParams.get('RequestID');
      if (requestid) {
        this.setState({ requestID: requestid });
        this.GetTaskDetailsItems(requestid);
        this.GetProjectNameDetailsItem(requestid);
        this.GetProjectManagerDetailsItem(requestid);
      } else {
        console.log("RequestID not found in URL parameters.");
      }
    } catch (error) {
      console.log("Error parsing URL Parameters: ", error);
    }
  }

  public async GetTaskDetailsItems(ID) {
    const tasks = await sp.web.lists.getByTitle("Project Task list").items.select(
      "ID",
      "TaskName",
      "Description",
      "StartDate",
      "EndDate",
      "ProjectManager/ID",
      "ProjectManager/ProjectManager",
      "Status",
      "Priority",
      "AssignedTo/Id",
      "AssignedTo/Title",
      "AssignedTo/EMail",
      "ProjectName/ID",
      "ProjectName/ProjectName",
      "RequestID/Id",
    ).expand("AssignedTo", "ProjectName", "ProjectManager", "RequestID").filter(`RequestID/Id eq ${ID}`).get().then((data) => {
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
            Priority: item.Priority ? item.Priority : "",
            AssignedTo: item.AssignedTo ? item.AssignedTo : "",
            ProjectNameId: item.ProjectName ? item.ProjectName.ID : "",
            ProjectName: item.ProjectName ? item.ProjectName.ProjectName : "",
            RequestID: item.RequestID ? item.RequestID.Id : "",
          });
        });
        this.setState({ TaskDetailsData: AllData, AllTaskListDetails: AllData });
        console.log(this.state.TaskDetailsData);
        console.log("Request ID :", this.state.requestID);
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
    if (this.state.TaskName.length == 0) {
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
        AssignedToId: { results: this.state.AssignedToID },
        RequestIDId: this.state.requestID
      }).catch((error) => {
        console.log("Can't add a new task", error);
      });

      this.GetTaskDetailsItems(this.state.requestID);
      this.setState({ TaskDetailsData: addtaskdetails });
      this.setState({ AddTaskDialogOpen: true });
      this.setState({ TaskFormSection1: true, TaskFormSection2: false, TaskFormSection3: false });

    }
  }

  public async GetEditTaskDetails(ID) {
    let EditTaskDetails = this.state.TaskDetailsData.filter((item) => {
      if (item.ID == ID) {
        return item;
      }
    });
    console.log(EditTaskDetails);
    this.setState({
      EditTaskName: EditTaskDetails[0].TaskName,
      EditDescription: EditTaskDetails[0].Description,
      EditStartDate: new Date(EditTaskDetails[0].StartDate),
      EditEndDate: new Date(EditTaskDetails[0].EndDate),
      EditProjectManagerID: EditTaskDetails[0].ProjectManagerId,
      // EditProjectManager: EditTaskDetails[0].ProjectManager,
      EditStatus: EditTaskDetails[0].Status,
      EditPriority: EditTaskDetails[0].Priority,
      EditAssignedTo: EditTaskDetails[0].AssignedTo.map(item => item.EMail),
      EditProjectNameID: EditTaskDetails[0].ProjectNameId,
      // EditProjectName: EditTaskDetails[0].ProjectName,
      // EditAssignedToID: EditTaskDetails[0].AssignedToID
    });

  }

  public async UpateTaskDetails(CurrentTaskDetailsID) {
    try {
      const updatetaskdetails: any = {
        TaskName: this.state.EditTaskName,
        Description: this.state.EditDescription,
        StartDate: this.state.EditStartDate,
        EndDate: this.state.EditEndDate,
        Status: this.state.EditStatus,
        Priority: this.state.EditPriority,
        ProjectNameId: this.state.EditProjectNameID,
        ProjectManagerId: this.state.EditProjectManagerID,
      };

      if (this.state.AssignedToID && this.state.AssignedToID.length > 0) {
        updatetaskdetails.AssignedToId = { results: this.state.AssignedToID };
      }

      const updatedetails = await sp.web.lists.getByTitle("Project Task list").items.getById(CurrentTaskDetailsID).update(updatetaskdetails);

      this.setState({ TaskDetailsData: updatedetails });

    } catch (error) {
      console.log("Error updating task details: ", error);
    }

    this.GetTaskDetailsItems(this.state.requestID);
    this.setState({ EditTaskDialogOpen: true });
    this.setState({ TaskFormSection1: true, TaskFormSection2: false, TaskFormSection3: false });

  }

  public async DeleteTaskDetails(DeleteTaskDetailsID) {
    const deletetaskdetails = await sp.web.lists.getByTitle("Project Task list").items.getById(DeleteTaskDetailsID).delete();
    this.setState({ TaskDetailsData: deletetaskdetails });
    this.setState({ DeleteTaskDialogOpen: true });
    this.GetTaskDetailsItems(this.state.requestID);
  }

  // publzic async GetProjectManagerDetailsItem() {
  //   try {
  //     const data = await sp.web.lists.getByTitle("Project Details").items
  //       .select("ID", "ProjectManager")
  //       .get();

  //     const uniqueNames = new Set<string>();
  //     const detailsData: { key: number, text: string }[] = [];

  //     data.forEach((d) => {
  //       if (d.ProjectManager && !uniqueNames.has(d.ProjectManager)) {
  //         uniqueNames.add(d.ProjectManager);
  //         detailsData.push({ key: d.ID, text: d.ProjectManager });
  //       }
  //     });

  //     this.setState({ ProjectManagerlist: detailsData });
  //   } catch (error) {
  //     console.error("Error fetching project manager details:", error);
  //   }
  // }

  public async GetProjectManagerDetailsItem(ID) {
    try {
      const data = await sp.web.lists.getByTitle("Project Details").items.select(
        "ID",
        "ProjectManager"
      ).getById(ID).get();

      const detailsItem = [];
      detailsItem.push({ key: data.ID, text: data.ProjectManager });
      console.log("Project Manager Details: ", detailsItem);
      this.setState({ ProjectManagerlist: detailsItem });

    } catch (error) {
      console.log(error);
    }
  }

  public async GetProjectNameDetailsItem(ID) {
    try {
      const data = await sp.web.lists.getByTitle("Project Details").items
        .select("ID", "ProjectName")
        .getById(ID)
        .get();

      const detailsData = [];
      detailsData.push({ key: data.ID, text: data.ProjectName });
      console.log("Project Name: ", detailsData);
      this.setState({ ProjectNamelist: detailsData });

    } catch (error) {
      console.log(error);
    }
  }

  public async GetTaskdetailsStatusItems() {
    const choieFieldName1 = "Status";
    const field1 = await sp.web.lists.getByTitle("Project Task list").fields.getByInternalNameOrTitle(choieFieldName1)();
    let status = [];
    field1["Choices"].forEach(function (dname, i) {
      status.push({ key: dname, text: dname });
    });
    this.setState({ Statuslist: status });

    const choiceFieldName2 = "Priority";
    const field2 = await sp.web.lists.getByTitle("Project Task list").fields.getByInternalNameOrTitle(choiceFieldName2)();
    let priority = [];
    field2["Choices"].forEach(function (dname, i) {
      priority.push({ key: dname, text: dname });
    });
    this.setState({ Prioritylist: priority });
  }

  public _getPeoplePickerItems = async (items: any[]) => {

    if (items.length > 0) {

      const memberNames = items.map(item => item.text);
      const memberIDs = items.map(item => item.id);
      this.setState({ AssignedTo: memberNames });
      this.setState({ AssignedToID: memberIDs });
    }
    else {
      this.setState({ AssignedTo: [] });
      this.setState({ AssignedToID: [] });
    }
  }

}



// <div className='ms-Grid-row'>
//             <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
//               <div className='Status-box'>
//                 {
//                   this.state.TaskDetailsData.length > 0 &&
//                   this.state.TaskDetailsData.map((item) => {
//                     return (
//                       item.Status == "To Do" ?
//                         <>
//                           <div className='Status-kit'>
//                             <div className='Status'>
//                               <h2 className='InProgress-Title'>To Do</h2>

//                               <div className='Inprogress-card Status-wrapper'>
//                                 <h3 className="card-title">{item.TaskName}</h3>
//                                 <p>{item.Description}</p>
//                                 {
//                                   item.Priority == "Low" ?
//                                     <>
//                                       <span className="badge-low">{item.Priority}</span>
//                                     </>
//                                     :
//                                     <>
//                                       {
//                                         item.Priority == "Medium" ?
//                                           <>
//                                             <span className="badge-medium">{item.Priority}</span>
//                                           </>
//                                           :
//                                           <>
//                                             {
//                                               item.Priority == "High" ?
//                                                 <>
//                                                   <span className="badge-high">{item.Priority}</span>
//                                                 </>
//                                                 :
//                                                 <></>
//                                             }
//                                           </>
//                                       }
//                                     </>
//                                 }

//                               </div>
//                             </div>

//                           </div>
//                         </>
//                         :
//                         <>
//                           {
//                             item.Status == "In Progress" ?
//                               <>
//                                 <div className='Status-Kit'>
//                                   <div className='Status'>
//                                     <h2 className='InProgress-Title'>In Progress</h2>

//                                     <div className='Inprogress-card Status-wrapper'>
//                                       <h3 className="card-title">{item.TaskName}</h3>
//                                       <p>{item.Description}</p>
//                                       {
//                                         item.Priority == "Low" ?
//                                           <>
//                                             <span className="badge-low">{item.Priority}</span>
//                                           </>
//                                           :
//                                           <>
//                                             {
//                                               item.Priority == "Medium" ?
//                                                 <>
//                                                   <span className="badge-medium">{item.Priority}</span>
//                                                 </>
//                                                 :
//                                                 <>
//                                                   {
//                                                     item.Priority == "High" ?
//                                                       <>
//                                                         <span className="badge-high">{item.Priority}</span>
//                                                       </>
//                                                       :
//                                                       <></>
//                                                   }
//                                                 </>
//                                             }
//                                           </>
//                                       }
//                                     </div>
//                                   </div>
//                                 </div>
//                               </>
//                               :
//                               <>

//                               </>
//                           }
//                         </>



//                     );
//                   })

//                 }
//               </div>
//             </div>
// </div>