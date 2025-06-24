import * as React from 'react';
import styles from './ProjectDetails.module.scss';
import { IProjectDetailsProps } from './IProjectDetailsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IIconProps } from 'office-ui-fabric-react';





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
      ProjectDetailsAddOpenDialog : true
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
        <h3>Project Details</h3>
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
      const addProjectdetails = await sp.web.lists.getByTitle("ProjectDetails").items.add({
        ProjectName : this.state.ProjectName,
        ProjectDescription : this.state.ProjectDescription,
        ProjectStartDate : this.state.ProjectStartDate,
        ProjectEndDate : this.state.ProjectEndDate,
        ProjectStatus : this.state.ProjectStatus,
        ProjectManager : this.state.ProjectManager,
        AssignedTo : this.state.AssignedTo,
        Attachments : this.state.Attachments
      })
    }
  }
}
