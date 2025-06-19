import * as React from 'react';
import styles from './ProjectDetails.module.scss';
import { IProjectDetailsProps } from './IProjectDetailsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IIconProps } from 'office-ui-fabric-react';





export interface IProjectDetailsState {

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
        
      </section>
    );
  }

  public async componenetDidMount() {
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
    )
  }
}
