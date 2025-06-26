import * as React from 'react';
import styles from './ProjectDetails.module.scss';
import { IProjectDetailsProps } from './IProjectDetailsProps';
import { escape, update } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { DatePicker, Dialog, IIconProps, PrimaryButton, SearchBox, TextField } from 'office-ui-fabric-react';
import { IItem, Item } from '@pnp/sp/items';
import { Attachments } from '@pnp/sp/attachments';




export interface IProjectDetailsState {
  ProjectDetails : any;
  ProjectName : any;
  ProjectDescription : any;
  ProjectStartDate : any;
  ProjectEndDate : any;
  ProjectStatus : any;
  ProjectManager : any;
  AssignedTo : any;
  Attachments : any;
  ProjectDetailsAddOpenDialog : boolean;
  RemoveAttachment : any;
  UploadDocuments : any;
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

  constructor(props : IProjectDetailsProps, state : IProjectDetailsState) {

    super(props);

    this.state = {
      ProjectDetails : "",
      ProjectName : "",
      ProjectDescription : "",
      ProjectStartDate : "",
      ProjectEndDate : "",
      ProjectStatus : [],
      ProjectManager : "",
      AssignedTo : "",
      Attachments : "",
      ProjectDetailsAddOpenDialog : true,
      RemoveAttachment : [],
      UploadDocuments : []
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
                  onSelectDate={(date: any) => this.setState({ ProjectStartDate : date })}
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

          </div>

        </Dialog>

      </section>
    );
  }

  public async componenetDidMount() {
    this.GetProjectDetails();
  }

  public async GetProjectDetails() {
    const projectdetails = await sp.web.lists.getByTitle("ProjectDetails").items.select(
      "ID",
      "ProjectName",
      "ProjectDescription",
      "ProjectStartDate",
      "ProjectEndDate",
      "ProjectStatus",
      "ProjectManager/Id",
      "ProjectManager/Title",
      "AssignedTo/Id",
      "AssignedTo/Title",
      "Attachments"
    ).expand("AssignedTo","ProjectManager").get().then((data) => {
      let AllData = [];
      console.log(data);
      console.log(projectdetails);

      if(data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID : item.Id ? item.Id : "",
            ProjectName : item.ProjectName ? item.ProjectName : "",
            ProjectDescription : item.ProjectDescription ? item.ProjectDescription : "",
            ProjectStartDate : item.ProjectStartDate ? item.ProjectStartDate : "",
            ProjectEndDate : item.ProjectEndDate ? item.ProjectEndDate : "",
            ProjectStatus : item.ProjectStatus ? item.ProjectStatus : "",
            ProjectManager : item.ProjectManager ? item.ProjectManager.Title : "",
            AssignedTo : item.AssignedTo ? item.AssignedTo.Title : "",
            Attachments : item.Attachments ? item.Attachments : ""
          });
        });
        this.setState({ ProjectDetails : AllData });
        console.log(this.state.ProjectDetails);
      }
    }).catch((error) => {
      console.log("Error fetching project details: ", error);
    });
  }


  public async AddProjectDetails() {
    if(this.state.ProjectName.length == 0) {
      alert("Please enter Project Details");
    }
    else {
      const addProjectdetails : any = await sp.web.lists.getByTitle("ProjectDetails").items.add({
        ProjectName : this.state.ProjectName,
        ProjectDescription : this.state.ProjectDescription,
        ProjectStartDate : this.state.ProjectStartDate,
        ProjectEndDate : this.state.ProjectEndDate,
        ProjectStatus : this.state.ProjectStatus,
        ProjectManager : this.state.ProjectManager,
        AssignedTo : this.state.AssignedTo,
        Attachments : this.state.Attachments
      })
      .catch((error) => {
        console.log(error);
      });

      for(let i = 0; i < this.state.RemoveAttachment.length; i++) {
        const file = this.state.RemoveAttachment[i];

        try {
          const item1: IItem = await sp.web.lists.getByTitle("ProjectDetails").items.getById(file.Id);
          await item1.attachmentFiles.getByName(file.FileName).delete();
        } catch (error) {
          console.log(`Error : ${file.FileName}`);
        }
      }
      this.setState({ RemoveAttachment : [] });

      for(let i = 0; i < this.state.UploadDocuments.length; i++) {
        const file = this.state.UploadDocuments[i];

        try {
          const item2: IItem = await sp.web.lists.getByTitle("ProjectDetails").items.getById(file.Id);
          await item2.attachmentFiles.add(file.FileName, file.FileContent);
        } catch(error) {
           console.log(`Error uploading file ${file.FileName}: `, error);
        }
      }
      this.GetProjectDetails();
      this.setState({ UploadDocuments : [] });
      this.setState({ ProjectDetails : addProjectdetails });
      this.setState({ ProjectDetailsAddOpenDialog : true });
    }
  }

  public async UploadAttachments(files, id: number, Title) {
    const updateProjectdetailsdoc = this.state.ProjectDetails.map(item => {
      if(item.Title === Title) {
        return {
          ...item,
          file: item.file ? [...item.file, ...files] : [...files],
          isfilechanged: true,
        };
      }
      else {
        return item; 
      }
    });

    const filteredData = this.state.ProjectDetails.filter(item => item.Title === Title);

    const uploadeddoc = this.state.UploadDocuments;

    const fileArray = [...files];
    fileArray.map(item => {
      uploadeddoc.push({ 
        Id : filteredData[0].Id,
        FileName : item.name,
        file : item,
        DocumentType : Title
      })
    })
  }

  public async RemoveUploadedDoc(id : number, file) {
    const fileToRemove = file;

    const updateddetails = this.state.ProjectDetails.map(item => {
      if(item.Id === id) {
        const files = Array.isArray(item.file) ? item.file : [item.file];
        const updatedoc = files.filter(f => f.name !== fileToRemove);
        return {
          ...item,
          file: updatedoc,
        };
      }
      return item;
    });

    const updatedoc = this.state.UploadDocuments;
    updatedoc.filter(f => f.FileName !== fileToRemove);

    this.setState({ ProjectDetails : updateddetails, UploadDocuments : updatedoc });
  }

  public async RemoveAttachments(id: number, fileName: string) {
    const updated = this.state.ProjectDetails.map(item => {
      if(item.Id === id) {
        const updatedFiles = item.Attachments.filter(f => f.FileName !== fileName);
        return {
          ...item,
          Attachments : updatedFiles,
        };
      }
      else {
        return item;
      }
    });

    let DeleteDocs = this.state.RemoveAttachment;
    DeleteDocs.push({
      FileName: fileName,
      Id : id
    });
    this.setState({ ProjectDetails : updated, RemoveAttachment : DeleteDocs });
  }
}
