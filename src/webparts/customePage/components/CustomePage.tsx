import * as React from 'react';
import styles from './CustomePage.module.scss';
import { ICustomePageProps } from './ICustomePageProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';
import * as moment from 'moment';
import { Icon } from 'office-ui-fabric-react';


export interface ICustomePageState {
  AllVirtualMapData : any;
  AllQuestions: any;
  AllGotQuestionData : any;
  ArrowButton: boolean; 
  ReadAnswer: any;
  Answer : any;
  Downbutton: boolean;
  SelectedId : any;
}

require("../assets/css/style.css");

export default class CustomePage extends React.Component<ICustomePageProps, ICustomePageState> {

  constructor(props: ICustomePageProps , state: ICustomePageState) {
    super(props);

    this.state = {
      AllVirtualMapData : "",
      AllQuestions : "",
      AllGotQuestionData : "",
      ArrowButton : true,
      ReadAnswer : "",
      Answer : "",
      Downbutton : true,
      SelectedId : ""
    };
  }

  public render(): React.ReactElement<ICustomePageProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
        <div className="customePage">
          <div className='ms-Grid'>
            <div className='ms-Grid-row'>
              <div className='ms-Grid-col'> 
                <div className='Virtual-board'>
                  <h2>VIRTUAL ONBOARDING ROADMAP</h2>
                  <div className='Roadmap-Resource'>
                      {
                        this.state.AllVirtualMapData.length > 0 && 
                          this.state.AllVirtualMapData.map((item) => {
                            return (
                              <>
                                <div className='ms-Grid-col'>
                                  
                                    {
                                      item.Title == "Step 1" ?
                                      <>
                                        <div className='Step1-Resource'>
                                            <div className="Steps-1">
                                                <h3>{item.Title}</h3>
                                                <p>{item.Description}</p>
                                              <div className='Steps1-Icons'>
                                                <Icon iconName='DietPlanNotebook' className='Step1-Icon'/>
                                              </div>
                                            </div>
                                            <div className="Step1-line"></div>
                                        </div>
                                        
                                      </>
                                      : 
                                      <>
                                        {
                                          item.Title == "Step 2" ?
                                          <>
                                            <div className='Step2-Resource'>
                                              <div className="Steps-2">
                                                  <h3>{item.Title}</h3>
                                                  <p>{item.Description}</p>
                                                <div className='Steps2-Icons'>
                                                  <Icon iconName='CompassNW' className='Step2-Icon'/>
                                                </div>
                                              </div>
                                              <div className="Step2-line"></div>
                                            </div>
                                            
                                          </>
                                          :
                                          <>
                                          {
                                            item.Title == "Step 3" ?
                                            <>
                                              <div className='Step3-Resource'>
                                                <div className="Steps-3">
                                                  <h3>{item.Title}</h3>
                                                  <p>{item.Description}</p>
                                                  <div className='Steps3-Icons'>
                                                    <Icon iconName='RedEye12' className='Step3-Icon'/>
                                                  </div>
                                                </div>
                                                <div className="Step3-line"></div>
                                              </div>
                                             
                                            </>
                                            :
                                            <>
                                            {
                                              item.Title == "Step 4" ?
                                              <>
                                                <div className='Step4-Resource'>
                                                  <div className="Steps-4">
                                                    <h3>{item.Title}</h3>
                                                    <p>{item.Description}</p>
                                                    <div className='Steps4-Icons'>
                                                      <Icon iconName='Education' className='Step4-Icon'/>
                                                    </div>
                                                  </div>
                                                  <div className="Step4-line"></div>
                                                </div>
                                                
                                              </>
                                              :
                                              <>
                                              {
                                                item.Title == "Step 5" ?
                                                <>
                                                  <div className='Step5-Resource'>
                                                    <div className="Steps-5">
                                                      <h3>{item.Title}</h3>
                                                      <p>{item.Description}</p>
                                                      <div className='Steps5-Icons'>
                                                        <Icon iconName='12PointStar' className='Step5-Icon'/>
                                                      </div>
                                                    </div>
                                                    <div className="Step5-line"></div>
                                                      <img src={require("../assets/Image/bxs--right-arrow.png")} className='Arrow-Img' />
                                                    {/* <Icon iconName='PlaySolid' className='Play-Arrow'/> */}
                                                  </div>
                                                  
                                                </>
                                                :
                                                <></>
                                              }
                                              </>
                                            }
                                            </>
                                          }
                                          </>
                                        }
                                      </>
                                    }
                                  </div>
                              </>
                            );
                          })
                      }
                  </div>

                </div>
              </div>
            </div>

          <div className='Question-Ans'>
            <div className='ms-Grid-row'>
              <div className='ms-Grid-col'> 
                  {
                    this.state.AllQuestions.length > 0 && 
                      this.state.AllQuestions.map((item) => {
                        return(
                          <>
                          <div className='ms-Grid-col'>
                            <div className='question'>
                              <div className='qustion-item'> 
                                <h3>
                                  {item.Question}<Icon className='arrow-down' iconName={this.state.ArrowButton == false && this.state.SelectedId == item.ID ?  "ChevronUp" : "ChevronDown"} onClick={()=> this.setState({ SelectedId : item.ID }, () => this.handleClick(item.ID))}></Icon>
                                </h3>
                              </div>
                              <div className='Questions'>
                                {
                                  (this.state.ArrowButton == false && this.state.SelectedId == item.ID) ? 
                                    <>
                                      <div dangerouslySetInnerHTML={{__html: item.Answer}} className='Answer-Item'/>
                                    </> :
                                    <></>
                                }                                      
                              </div>
                            </div>
                          </div>
                          </>
                        );
                      })
                  }
              </div>
            </div>
          </div>

          </div>
        </div>
    );
  }

  public async componentDidMount() {
    this.GetVirtualMap();
    this.GetQuestion();
  }

  public async GetVirtualMap() {
    const map = await sp.web.lists.getByTitle("Virtual RoadMap").items.select(
      "ID",
      "Title",
      "Description"
    ).get().then((data) => {
      let AllData = [];
      console.log(data);
      console.log(map);

      if(data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID: item.Id ? item.Id : "",
            Title: item.Title ? item.Title : "",
            Description : item.Description ? item.Description : ""
          });
        });
        this.setState({ AllVirtualMapData : AllData });
        console.log(this.state.AllVirtualMapData);
      }
    }).catch((error) => {
      console.log("Error Retrived" , error);
    });
  }

  public async GetQuestion() {
    const question = await sp.web.lists.getByTitle("Questions").items.select(
      "ID",
      "Question",
      "Answer"
    ).get().then((data) => {
      let AllData = [];
      console.log(data);
      console.log(question);
      if(data.length > 0) {
        data.forEach((item) => {
          AllData.push({
            ID : item.Id ? item.Id : "",
            Question : item.Question ? item.Question : "",
            Answer : item.Answer ? item.Answer : ""
          });
        });
        this.setState({ AllQuestions : AllData });
      }
    }).catch((error) => {
      console.log("Error Retrived" , error);
    });
  }

  public handleClick = (id) => {
    // Conditional logic to set different state values
    if (this.state.ArrowButton === false) {
      this.setState({ArrowButton:true});
      this.setState({SelectedId:id});
    } else if(this.state.ArrowButton === true && this.state.SelectedId == id){
      this.setState({ArrowButton:true});
      this.setState({ArrowButton:false});
      this.setState({SelectedId:id});
    }
    else {
      this.setState({ArrowButton:false});
      this.setState({SelectedId:id});
    }
    console.log(this.state.ArrowButton);
    console.log(this.state.SelectedId);
  }

}
