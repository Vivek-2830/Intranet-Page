import * as React from 'react';
import styles from './Contact.module.scss';
import { IContactProps } from './IContactProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';
import * as moment from 'moment';
import { Dialog, Icon, PrimaryButton, TextField } from 'office-ui-fabric-react';

export interface IContactState {
  ContactData : any;
  Questions: any;
  EmailAddress : any;
  AddFAQQuestion : any;
  AddFAQDialog : boolean;
  ContactDialog: boolean;
  ContactTwoData : any;
}

require("../assets/css/style.css");

const AddFaqQuestionsDialogContentProps = {
  title : "Frequently Asked Questions(FAQ)"
};

const ContactDialogContentProps = {
};

const addmodelProps = {
  className: "Add-Dialog"
};

const contactmodelProps = {
  className: "Contact-Dialog"
};

export default class Contact extends React.Component<IContactProps, IContactState> {

  constructor(props : IContactProps , state: IContactState) {
    super(props);

    this.state = {
      ContactData : "",
      Questions: "",
      EmailAddress : "",
      AddFAQQuestion : "",
      AddFAQDialog : true,
      ContactDialog : true,
      ContactTwoData : ""
    };
  }


  public render(): React.ReactElement<IContactProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      
        <div className="contact">
          <div className='ms-Grid'>
            <div className='ms-Grid-row'>
              <div className='ms-Grid-col'>
                <div className="Contact-Text">
                  <h2>CONTACTS</h2>
                  {
                    this.state.ContactData.length > 0 && 
                      this.state.ContactData.map((item) => {
                        return(
                          <>
                            <div className='ms-Grid-col'>
                              <div className='Contacts'>
                                <div className='Contact-Image'>
                                  <img src={item.Image} />
                                </div>
                                <div className='Contact-Details'>
                                  <h4>{item.Title}</h4>
                                  <p>{item.Description}</p>
                                </div>
                              </div>
                            </div>

                            
                          </>
                        );
                      })
                  }
                  <hr className='line' />

                  <div className='More-Contact'>
                    <Icon iconName='ChevronDownMed' onClick={() => this.setState({ ContactDialog : false} , () => this.GetContactTwo())} />
                  </div>

                  <div className='FAQ-Question'>
                    <img src="https://miro.medium.com/v2/resize:fit:915/0*FdLLSjLPudGd-Pt5" alt="FAQ Questions" onClick={() => this.setState({ AddFAQDialog : false })}/>
                    <h2>FAQS</h2>
                  </div>

                </div>
              </div>
            </div>
          </div>

          <Dialog
            hidden={this.state.ContactDialog}
             onDismiss={() =>
              this.setState({
                ContactDialog : true
              })
             }
            dialogContentProps={ContactDialogContentProps}
            modalProps={contactmodelProps}
            minWidth={400}
          >
            {
              this.state.ContactTwoData.length > 0 && 
                this.state.ContactTwoData.map((item) => {
                  return(
                    <>
                      <div className='ms-Grid-col'>
                        <div className='Contacts-Two'>
                          <div className='ContactTwo-Image'>
                            <img src={item.Image} />
                          </div>
                          <div className='ContactTwo-Details'>
                            <h4>{item.Title}</h4>
                            <p>{item.Description}</p>
                          </div>
                        </div>
                      </div>
                    </>
                  );
                })
              }
          </Dialog>

          <Dialog
            hidden={this.state.AddFAQDialog}
            onDismiss={() => 
              this.setState({
                AddFAQDialog : true,
                Questions : "",
                EmailAddress : ""
              })
            }
            dialogContentProps={AddFaqQuestionsDialogContentProps}
            modalProps={addmodelProps}
            minWidth={500}
          >
              <div>
                <div className='ms-Grid-row'>
                  <div className='ms-Grid-col  ms-sm12 ms-md12 ms-lg12'>
                    <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                      <TextField
                        label='Ask Your Questions?'
                        name="Questions"
                        type='Text'
                        placeholder='Please Enter Your Asking Question?'
                        required={true}
                        onChange={(value) => 
                          this.setState({Questions : value.target["value"]})
                        }
                        value={this.state.Questions}
                      />
                    </div>

                    <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'>
                      <TextField
                        label='Email Address'
                        name="Questions"
                        type='Text'
                        placeholder='Please Enter Your Email Address'
                        required={true}
                        onChange={(value) => 
                          this.setState({EmailAddress : value.target["value"]})
                        }
                        value={this.state.EmailAddress}
                      />
                    </div>

                  </div>
                </div>

                <div className='ms-Grid-row'>
                  <div className='Submit-Question'>
                    <PrimaryButton
                      type='Submit'
                      text="Submit"
                      onClick={() => this.AddFAQQuestions()}
                    />
                  </div>
                </div>

              </div>

          </Dialog>

        </div>
    );
  }

  public async componentDidMount() {
    this.GetContacts();
    this.GetFAQQuestions();
    this.GetContactTwo();
  }

  public async GetContacts() {
    const contacts = await sp.web.lists.getByTitle("Contacts").items.select(
      "ID",
      "Title",
      "Description",
      "Image"
    ).get().then((data) => {
      let AllData = [];
      console.log(data);
      console.log(contacts);
      if(data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID : item.Id ? item.Id : "",
            Title: item.Title ? item.Title : "",
            Description: item.Description ? item.Description : "",
            Image: item.Image ? JSON.parse(item.Image).serverRelativeUrl : ""
          });
        });
        this.setState({ ContactData : AllData });
      }
    }).catch((error) => {
      console.log("Error Retrived" , error);
    });
  }

  public async GetFAQQuestions() {
    const faqque = await sp.web.lists.getByTitle("FAQ Questions").items.select(
      "ID",
      "Questions",
      "EmailAddress"
    ).get().then((data) => {
      let AllData = [];
      console.log(data);
      console.log(faqque);

      if(data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID : item.Id ? item.Id : "",
            Questions : item.Questions ? item.Questions : "",
            EmailAddress : item.EmailAddress ? item.EmailAddress : ""
          });
        });
        this.setState({ AddFAQQuestion : AllData });
        console.log(this.state.AddFAQQuestion);
      }
    }).catch((error) => {
      console.log("Error Retrived" , error);
    });
  }

  public async GetContactTwo() {
    const contacttow = await sp.web.lists.getByTitle("Contacts 2").items.select(
      "Title",
      "Description",
      "Image"
    ).get().then((data) => {
      let AllData = [];
      console.log(data);
      console.log(contacttow);
      if(data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID : item.Id ? item.Id : "",
            Title: item.Title ? item.Title : "",
            Description: item.Description ? item.Description : "",
            Image: item.Image ? JSON.parse(item.Image).serverRelativeUrl : ""
          });
        });
        this.setState({ ContactTwoData : AllData });
        // this.setState({ ContactDialog : false });
      }
    }).catch((error) => {
      console.log("Error Retrived" , error);
    });
  }
  

  public async AddFAQQuestions() {
    if(this.state.Questions.length == 0 || this.state.EmailAddress.length == 0) {
      alert("Please Complete the Details");
    } else {
      const addfaqquestions = await sp.web.lists.getByTitle("FAQ Questions").items.add({
        Questions : this.state.Questions,
        EmailAddress : this.state.EmailAddress
      }).catch((error) => {
        console.log(error);
      });
      this.setState({ AddFAQQuestion : addfaqquestions });
      this.GetFAQQuestions();
    }
  }

}
