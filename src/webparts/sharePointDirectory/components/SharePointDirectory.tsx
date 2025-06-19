import * as React from 'react';
import styles from './SharePointDirectory.module.scss';
import { ISharePointDirectoryProps } from './ISharePointDirectoryProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp';
import { DefaultButton, Dialog, Dropdown, IIconProps, PrimaryButton, TextField } from 'office-ui-fabric-react';

export interface ISharePointDirectoryState {
  AddFormDialog : boolean;
  Citylist : any;
  AgeCategorylist : any;
  SharePointDirectoryData : any;
  Title : any;
  Email : any;
  Description : any;
  AgeCategory : any;
  City : any;
}

const addIcon: IIconProps = { iconName: 'Add' };

const SendIcon: IIconProps = { iconName: 'Send' };

const CancelIcon: IIconProps = { iconName: 'Cancel' };

const AddFormDesignDialogContentProps = {
  title: "Add Checklist Details",
};

const addmodelProps = {
  className: "Add-Dialog"
};

require("../assets/css/style.css");
require("../assets/css/fabric.min.css");

export default class SharePointDirectory extends React.Component<ISharePointDirectoryProps, ISharePointDirectoryState> {

  constructor(props: ISharePointDirectoryProps, state: ISharePointDirectoryState) {
    
    super(props);

    this.state = {
      AddFormDialog : true,
      Citylist : [],
      AgeCategorylist : [],
      SharePointDirectoryData : "",
      Title : "",
      Email : "",
      Description : "",
      AgeCategory : [],
      City : [],
    };

  }

  public render(): React.ReactElement<ISharePointDirectoryProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
        <div className="sharePointDirectory">
          <div className='ms-Grid'>
            <div className='ms-Grid-row'>
              <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg4'>
                <h3>SharePoint Modern Form</h3>
              </div>

              <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg5 Add-Form'>
                <div className='Add-Form-container'>
                  <PrimaryButton iconProps={addIcon} text="Add Details" onClick={() => this.setState({ AddFormDialog : false })}
                  />
                </div>
              </div>

            </div>
          </div>

          <div className='ms-Grid-row'>
            <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
              { 
                this.state.SharePointDirectoryData.length > 0 &&
                  this.state.SharePointDirectoryData.map((item) => {
                    return (
                      <div className='ms-Grid-col ms-sm6 ms-md3 ms-lg3'>
                        <div className='Card'>
                          <div className='Card-Header'>
                            <h3>{item.Title}</h3>
                          </div>
                          <div>
                            {item.Description}
                          </div>
                          <div>
                            {item.Email}
                          </div>
                          <div>
                            {item.City}
                          </div>
                          <div>
                            {item.AgeCategory}
                          </div> 
                        </div>
                      </div>
                    );
                  })
              }
            </div>
          </div>

          <Dialog
            hidden={this.state.AddFormDialog} 
            onDismiss={() =>
              this.setState({
                AddFormDialog: true
              })
            }
            dialogContentProps={AddFormDesignDialogContentProps}
            modalProps={addmodelProps}
            minWidth={500}
          >
            
            <div className='ms-Grid-row'>
              <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                <div className='Add-Form-Details'>
                  <TextField
                    label='Title'
                    type='Text'
                    required={true}
                    onChange={(value) =>
                      this.setState({ Title : value.target["value"] })
                    }
                  />
                </div>
              </div>

              <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                <div className='Add-Form-Details'>
                  <TextField
                    label='Description'
                    type="Text"
                    multiline={true}
                    rows={3}
                    required={true}
                    onChange={(value) =>
                      this.setState({ Description : value.target["value"] })
                    }
                  />
                </div>
              </div>

              <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg6'>
                <div className='Add-Form-Details'>
                  <Dropdown
                    options={this.state.Citylist}
                    label='Select City'
                    required
                    placeholder="Enter Your City"
                    onChange={(e, option, text) =>
                      this.setState({ City : option.text })
                    }
                   />
                </div>
              </div>

              <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg6'>
                <div className='Add-Form-Details'>
                  <Dropdown
                    options={this.state.AgeCategorylist}
                    label='Select Age'
                    required
                    placeholder='Enter Your Age'
                    onChange={(e, option, text) => 
                      this.setState({ AgeCategory : option.text})
                    }
                  />
                </div>
              </div> 

              <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                <div className='Add-Form-Details'>
                  <TextField
                    label='Email'
                    type='Email'
                    required={true}
                    onChange={(value) =>
                      this.setState({ Email : value.target["value"] })
                    }
                   />
                </div>
              </div>   

            </div>

            <div className='ms-Grid-row'>
              <div className='Add-Directoy'>
                <div className='ms-Grid-col Add-Form'>
                  <PrimaryButton
                    iconProps={SendIcon}
                    text="Submit"
                    onClick={() =>
                      this.AddSharePointDirectoryDetails()} 
                  />
                
                  <div className='ms-Grid-col Add-Form-Cancel'>
                    <DefaultButton
                      iconProps={CancelIcon}
                      text="Cancel"
                      onClick={() => this.setState({ AddFormDialog: true })} 
                    />
                  </div>
                </div>
              </div>
            </div>
          </Dialog>

        </div>
      
    );
  }

  public async componentDidMount() {
    this.GetSharePointDirectory();
    this.GetSharePoinDirectoryChoiceItems();
  }

  public async GetSharePointDirectory() {
    const directory = await sp.web.lists.getByTitle("SharePoint Directory").items.select(
      "ID",
      "Title",
      "Description",
      "City",
      "AgeCategory",
      "Email"
    ).get().then((data) => {
      let AllData = [];
      console.log(data);
      console.log(directory);
      if(data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID: item.Id ? item.Id : "",
            Title : item.Title ? item.Title : "",
            Description : item.Description ? item.Description : "",
            City : item.City ? item.City : "",
            AgeCategory : item.AgeCategory ? item.AgeCategory : "",
            Email : item.Email ? item.Email : ""
          });
        });
        this.setState({ SharePointDirectoryData : AllData });
        console.log(this.state.SharePointDirectoryData);
      }
    }).catch((error) => {
      console.log(error);
    });
  }

  public async AddSharePointDirectoryDetails() {
    if(this.state.Title.length == 0 ) {
      alert("Please Complete the Details..!!!");
    } else {
      const addDirectory = await sp.web.lists.getByTitle("SharePoint Directory").items.add({
        Title : this.state.Title,
        Description : this.state.Description,
        City : this.state.City,
        AgeCategory : this.state.AgeCategory,
        Email : this.state.Email
      })
      .catch((error) => {
        console.log(error);
      });
      
      this.setState({ AddFormDialog : true });
      this.setState({ SharePointDirectoryData : addDirectory });
      this.GetSharePointDirectory();
    }
  }

  public async GetSharePoinDirectoryChoiceItems() {
    const ChoiceFieldName1 = "City";
    const filed1 = await sp.web.lists.getByTitle("SharePoint Directory").fields.getByInternalNameOrTitle(ChoiceFieldName1)();
    let citylist = [];
    filed1["Choices"].forEach(function (dname , i) {
      citylist.push({ key : dname, text : dname });
    });
    console.log(filed1);
    this.setState({ Citylist : citylist });

    const ChoiceFieldName2 = "AgeCategory";
    const filed2 = await sp.web.lists.getByTitle("SharePoint Directory").fields.getByInternalNameOrTitle(ChoiceFieldName2)();
    let ageCategorylist = [];
    filed2["Choices"].forEach(function (dname , i) {
      ageCategorylist.push({ key : dname, text : dname });
    });
    console.log(filed2);
    this.setState({ AgeCategorylist : ageCategorylist });
  }

}
