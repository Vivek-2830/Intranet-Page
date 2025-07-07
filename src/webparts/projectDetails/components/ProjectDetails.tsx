import * as React from 'react';
import styles from './ProjectDetails.module.scss';
import { IProjectDetailsProps } from './IProjectDetailsProps';
import { escape, update } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { DatePicker, DefaultButton, DetailsList, Dialog, Dropdown, IColumn, Icon, IIconProps, Label, PrimaryButton, SearchBox, TextField } from 'office-ui-fabric-react';
import { IItem, Item } from '@pnp/sp/items';
import { Attachments, IAttachmentInfo } from '@pnp/sp/attachments';
import * as moment from 'moment';


export interface IProjectDetailsState {
  ProjectDetails: any;
  ProjectName: any;
  ProjectDescription: any;
  ProjectStartDate: any;
  ProjectEndDate: any;
  ProjectStatus: any;
  ProjectStatuslist: any;
  ProjectManager: any;
  AssignedTo: any;
  AssignedToID: any;
  Attachments: any;
  ProjectDetailsAddOpenDialog: boolean;
  RemoveAttachment: any;
  UploadDocuments: any;
  ProjectDocuments: any;
  GetAllProjectDocuments: any;
  TempId: number;
  DeleteDocuments: any;
  EditProjectDetails: any;
  EditProjectName: any;
  EditProjectDescription: any;
  EditProjectStartDate: any;
  EditProjectEndDate: any;
  EditProjectStatus: any;
  EditProjectStatuslist: any;
  EditProjectManager: any;
  EditAssignedTo: any;
  EditAssignedToID: any;
  EditAttachments: any;
}

const addIcon: IIconProps = { iconName: 'Add' };

const SendIcon: IIconProps = { iconName: 'Send' };

const CancelIcon: IIconProps = { iconName: 'Cancel' };

const DeleteIcon: IIconProps = { iconName: 'Delete' };

const TextDocumentEdit: IIconProps = { iconName: 'TextDocumentEdit' };

const AddProjectDetailsDialogContentProps = {
  title: "Add Project Details",
};

const ReadProjectDetailsDialogContentProps = {
  title: "Read Project Details"
};

const UpdateProjectDetailsDialogContentProps = {
  title: "Update Project Details"
};

const DeleteProjectDetailsFilterDialogContentProps = {
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

export default class ProjectDetails extends React.Component<IProjectDetailsProps, IProjectDetailsState> {

  constructor(props: IProjectDetailsProps, state: IProjectDetailsState) {

    super(props);

    this.state = {
      ProjectDetails: "",
      ProjectName: "",
      ProjectDescription: "",
      ProjectStartDate: "",
      ProjectEndDate: "",
      ProjectStatus: [],
      ProjectStatuslist: [],
      ProjectManager: "",
      AssignedTo: "",
      AssignedToID: "",
      Attachments: "",
      ProjectDetailsAddOpenDialog: true,
      RemoveAttachment: [],
      UploadDocuments: [],
      ProjectDocuments: [],
      GetAllProjectDocuments: [],
      TempId: 0,
      DeleteDocuments: [],
      EditProjectDetails: "",
      EditProjectName: "",
      EditProjectDescription: "",
      EditProjectStartDate: "",
      EditProjectEndDate: "",
      EditProjectStatus: [],
      EditProjectStatuslist: [],
      EditProjectManager: "",
      EditAssignedTo: "",
      EditAssignedToID: "",
      EditAttachments: "",
    };

  }

  public render(): React.ReactElement<IProjectDetailsProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    const columns: IColumn[] = [
      {
        key: "ProjectName",
        name: "Project Name",
        fieldName: "ProjectName",
        minWidth: 100,
        maxWidth: 200,
        isResizable: false,
      },
      {
        key: "ProjectDescription",
        name: "Project Description",
        fieldName: "ProjectDescription",
        minWidth: 100,
        maxWidth: 200,
        isResizable: false,
      },
      {
        key: "ProjectStartDate",
        name: "ProjectStartDate",
        fieldName: "ProjectStartDate",
        minWidth: 100,
        maxWidth: 200,
        isResizable: false,
        onRender: (item) => {
          return <span>{moment(new Date(item.ProjectStartDate)).format("DD-MM-YYYY")}</span>
        }
      },
      {
        key: "ProjectManager",
        name: "Project Manager",
        fieldName: "ProjectManager",
        minWidth: 100,
        maxWidth: 200,
        isResizable: false,
      },
      {
        key: "AssignedTo",
        name: "AssignedTo",
        fieldName: "AssignedTo",
        minWidth: 100,
        maxWidth: 200,
        isResizable: false,
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
                      <Icon className='Edit-Icon' iconName="Edit" ></Icon>
                    </div>

                    <div className='Delete-Icon'>
                      <Icon className='icon' iconName="Delete" ></Icon>
                    </div>

                  </div>
                </div>
              </div>
            </div>

          )
        }
      }
    ]

    return (
      <section id="projectDetails">

        <div className='ms-Grid'>

          <div className='Header-Title'>
            <h3>Project Details</h3>
          </div>

          <div className='ms-Grid-row'>
            <div className='filedGroup'>

              <div className='ms-Grid-col ms-sm35 ms-md4 ms-lg2'>
                <SearchBox
                  placeholder="Search"
                  className='new-Search'

                />
              </div>

              <div className='ms-Grid-col ms-sm1 ms-md1 ms-lg10 Add-Projects'>
                <div className='Add-ProjectDetails'>
                  <PrimaryButton
                    iconProps={addIcon}
                    text="Add Project"
                    onClick={() => this.setState({ ProjectDetailsAddOpenDialog: false })}
                  />
                </div>
              </div>

            </div>
          </div>


          <Dialog
            hidden={this.state.ProjectDetailsAddOpenDialog}
            onDismiss={() =>
              this.setState({
                ProjectDetailsAddOpenDialog: true,
                ProjectName: "",
                ProjectDescription: "",
                ProjectStartDate: "",
                ProjectEndDate: "",
                ProjectStatus: [],
                ProjectManager: "",
                AssignedTo: "",
                Attachments: ""
              })
            }
            dialogContentProps={AddProjectDetailsDialogContentProps}
            modalProps={addmodelProps}
            minWidth={500}
          >
            <div className="ms-Grid-row">

              <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                <div className="Add-ProjectName">
                  <TextField
                    label="Project Name"
                    type="text"
                    required={true}
                    onChange={(value) =>
                      this.setState({ ProjectName: value.target["value"] })
                    }
                  />
                </div>
              </div>

              <div className="ms-Grid-col ms-sm12 ms-md6 ms-lg6">
                <div className='Add-StartDate'>
                  <DatePicker
                    label='Start Date'
                    allowTextInput={false}
                    value={this.state.ProjectStartDate ? this.state.ProjectStartDate : null}
                    onSelectDate={(date: any) => this.setState({ ProjectStartDate: date })}
                    aria-label="Select a Date" placeholder='Select a Project Start Date' isRequired
                  />
                </div>
              </div>

              <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                <div className='Add-ProjectDescription'>
                  <TextField
                    label="Project Description"
                    type="text"
                    multiline rows={3}
                    required={true}
                    onChange={(value) =>
                      this.setState({ ProjectDescription: value.target["value"] })
                    }
                  />
                </div>
              </div>

              <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                <div className='Add-EndDate'>
                  <DatePicker
                    label='End Date'
                    allowTextInput={false}
                    value={this.state.ProjectEndDate ? this.state.ProjectEndDate : null}
                    onSelectDate={(date: any) => this.setState({ ProjectEndDate: date })}
                    aria-label="Select a Date" placeholder='Select a Project End Date' isRequired
                  />
                </div>
              </div>

              <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                <div className='Add-ProjectStatus'>
                  <Dropdown
                    options={this.state.ProjectStatuslist}
                    label="Project Status"
                    required
                    placeholder="Select Project Status"
                    onChange={(e, option, text) =>
                      this.setState({ ProjectStatus: option.text })
                    }
                  />
                </div>
              </div>

              <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                <div className='Add-ProjectManager'>
                  <TextField
                    label="Project Manager"
                    type="text"
                    required={true}
                    onChange={(value) =>
                      this.setState({ ProjectManager: value.target["value"] })
                    }
                  />
                </div>
              </div>

              <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                <div className='Add-AssignedTo'>
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
                <div className="Add-Attachments">
                  <label className='Attachmentlabel'>Attachments</label>
                </div>
                <label className='Attachmentlabel' htmlFor='Project Document'>Choose files</label>
                <input style={{ display: 'none' }} id="Project Document" type="file" multiple onChange={(e) => this.GetUploadAttachments(e.target.files, "Attachments", "")}></input>
                <div>
                  {
                    this.state.UploadDocuments.length > 0 && (
                      this.state.UploadDocuments.map((item) => {
                        return (
                          <>
                            {
                              item.text == "Attachments" ?
                                <>
                                  <div>
                                    <Label>{item.key.name}</Label>
                                    <Icon iconName='Cancel' onClick={(e) => this.RemoveAttachments(item.TempId, item.ID, item.key.name)}></Icon>
                                  </div>
                                </> : <></>
                            }
                          </>
                        )
                      })
                    )
                  }
                </div>
              </div>

            </div>

            <div className='ms-Grid-row'>
              <div className='Add-Projects'>

                <div className='ms-Grid-col Add-Submit'>
                  <PrimaryButton
                    iconProps={SendIcon}
                    text="Submit"
                    onClick={() => this.AddProjectDetails()}
                  />
                </div>

                <div className='ms-Grid-col Cancel-Project'>
                  <DefaultButton
                    iconProps={CancelIcon}
                    text="Cancel"
                    onClick={() =>
                      this.setState({ ProjectDetailsAddOpenDialog: true })
                    }
                  />
                </div>

              </div>
            </div>

          </Dialog>

          <div className='ms-Grid'>
            <DetailsList
              className='ProjectDetails-List'
              items={this.state.ProjectDetails}
              columns={columns}
              setKey="set"
              layoutMode={1}
              isHeaderVisible={true}
              selectionMode={0}
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
    this.GetProjectDetails();
    this.GetProjectDetailsItem();
  }

  public async GetProjectDetails() {
    const projectdetails = await sp.web.lists.getByTitle("Project Details").items.select(
      "ID",
      "ProjectName",
      "ProjectDescription",
      "ProjectStartDate",
      "ProjectEndDate",
      "ProjectStatus",
      "ProjectManager",
      "AssignedTo/Id",
      "AssignedTo/Title",
      "Attachments"
    ).expand("AssignedTo").get().then((data) => {
      let AllData = [];
      console.log(data);
      console.log(projectdetails);

      if (data.length > 0) {
        data.forEach((item, i) => {

          // const item1: IItem = sp.web.lists.getByTitle("Project Details").items.getById(item.Id);
          // const info: IAttachmentInfo[] = await item1.attachmentFiles();

          AllData.push({
            ID: item.Id ? item.Id : "",
            ProjectName: item.ProjectName ? item.ProjectName : "",
            ProjectDescription: item.ProjectDescription ? item.ProjectDescription : "",
            ProjectStartDate: item.ProjectStartDate ? item.ProjectStartDate : "",
            ProjectEndDate: item.ProjectEndDate ? item.ProjectEndDate : "",
            ProjectStatus: item.ProjectStatus ? item.ProjectStatus : "",
            ProjectManager: item.ProjectManager ? item.ProjectManager : "",
            AssignedTo: item.AssignedTo ? item.AssignedTo.Title : "",
            Attachments: item.Attachments ? item.Attachments : "",
            isfilechanged: false,
          });
        });
        this.setState({ ProjectDetails: AllData });
        console.log(this.state.ProjectDetails);
      }
    }).catch((error) => {
      console.log("Error fetching project details: ", error);
    });
  }

  public async AddProjectDetails() {
    if (this.state.ProjectName.length == 0) {
      alert("Please enter Project Details");
    }
    else {
      const addProjectdetails: any = await sp.web.lists.getByTitle("Project Details").items.add({
        ProjectName: this.state.ProjectName,
        ProjectDescription: this.state.ProjectDescription,
        ProjectStartDate: this.state.ProjectStartDate,
        ProjectEndDate: this.state.ProjectEndDate,
        ProjectStatus: this.state.ProjectStatus,
        ProjectManager: this.state.ProjectManager,
        AssignedTo: this.state.AssignedTo,
        Attachments: this.state.Attachments
      })
        .catch((error) => {
          console.log(error);
        });

      for (let i = 0; i < this.state.RemoveAttachment.length; i++) {
        const file = this.state.RemoveAttachment[i];

        try {
          const item1: IItem = await sp.web.lists.getByTitle("Project Details").items.getById(file.Id);
          await item1.attachmentFiles.getByName(file.FileName).delete();
        } catch (error) {
          console.log(`Error : ${file.FileName}`);
        }
      }
      this.setState({ RemoveAttachment: [] });

      for (let i = 0; i < this.state.UploadDocuments.length; i++) {
        const file = this.state.UploadDocuments[i];

        try {
          const item2: IItem = await sp.web.lists.getByTitle("Project Details").items.getById(file.Id);
          await item2.attachmentFiles.add(file.FileName, file.FileContent);
        } catch (error) {
          console.log(`Error uploading file ${file.FileName}: `, error);
        }
      }
      this.GetProjectDetails();
      this.setState({ UploadDocuments: [] });
      this.setState({ ProjectDetails: addProjectdetails });
      this.setState({ ProjectDetailsAddOpenDialog: true });
    }
  }

  public async GetUploadAttachments(files, Doctype, Title) {
    let ProjectDocumentslist = this.state.UploadDocuments;

    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      ProjectDocumentslist.push({ key: file, text: Doctype });
      ProjectDocumentslist.push({
        ID: "",
        Filename: file.name,
        DocumentType: Title,
        TempId: this.state.TempId + (i + 1),
      });
      let test = this.state.TempId + (i + 1);
      this.setState({ TempId: test });
    }
    this.setState({ UploadDocuments: ProjectDocumentslist });
    console.log(this.state.UploadDocuments);
  }

  public RemoveAttachments(tempid, Id, filename) {
    var array = this.state.GetAllProjectDocuments;
    var array2 = this.state.UploadDocuments;

    var index = array.findIndex(x => x.TempId === tempid);
    var index2 = array2.findIndex(x => x.key.name === filename);

    if (index !== -1) {
      array.splice(index, 1);
      this.setState({ GetAllProjectDocuments: array });
    }

    if (index2 !== -1) {
      array2.splice(index2, 1);
      this.setState({ UploadDocuments: array2 });
    }

    if (Id) {
      let deletedocuments = this.state.RemoveAttachment;

      deletedocuments.push(
        {
          ID: Id
        }
      );
      this.setState({ RemoveAttachment: deletedocuments });
    }

    console.log(this.state.RemoveAttachment);
    console.log(this.state.UploadDocuments);
    console.log(this.state.GetAllProjectDocuments);
  }

  public async GetProjectDetailsItem() {
    const choiceFieldName1 = "Project Status";
    const field1 = await sp.web.lists.getByTitle("Project Details").fields.getByInternalNameOrTitle(choiceFieldName1)();
    let projectstatuslist = [];
    field1["Choices"].forEach(function (dname, i) {
      projectstatuslist.push({ key: dname, text: dname });
    });
    this.setState({ ProjectStatuslist: projectstatuslist });
    console.log(this.state.ProjectStatuslist);
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

  public async GetEditProjectDetails(ID) {
    let EditProjectdetails = this.state.ProjectDetails.filter((item) => {
      if(item.ID == ID) {
        return item;
      }
    });
    console.log(EditProjectdetails);
    this.setState({ 
      EditProjectName: EditProjectdetails[0].ProjectName,
      EditProjectDescription: EditProjectdetails[0].ProjectDescription,
      EditProjectStartDate : EditProjectdetails[0].ProjectStartDate,
      EditProjectEndDate : EditProjectdetails[0].ProjectEndDate,
      EditProjectStatus: EditProjectdetails[0].ProjectStatus,
      EditProjectManager: EditProjectdetails[0].ProjectManager,
      EditAssignedTo: EditProjectdetails[0].AssignedTo,
      // EditAssignedToID: EditProjectdetails[0].AssignedToID,
      EditAttachments: EditProjectdetails[0].Attachments,
    });
  }

  public async UpdateProjectDetails(CurrentProjectDetailsID) {
    const updatedetails = await sp.web.lists.getByTitle("Project Details").items.getById(CurrentProjectDetailsID).update({
      ProjectName : this.state.EditProjectName,
      ProjectDescription : this.state.EditProjectDescription,
      ProjectStartDate : this.state.EditProjectStartDate,
      ProjectEndDate : this.state.EditProjectEndDate,
      ProjectStatus: this.state.EditProjectStatus,
      ProjectManager: this.state.EditProjectManager,
      AssignedTo: this.state.EditAssignedTo,
    }).catch((error) => {
      console.log(error);
    });

    

  }

}
