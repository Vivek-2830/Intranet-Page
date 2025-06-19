import * as React from 'react';
import styles from './SharepointTemplate.module.scss';
import { ISharepointTemplateProps } from './ISharepointTemplateProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';
import * as moment from 'moment';
import { Icon } from 'office-ui-fabric-react';

export interface ISharepointTemplateState {

}

require("../assets/css/style.css");

export default class SharepointTemplate extends React.Component<ISharepointTemplateProps, ISharepointTemplateState> {

  constructor(props: ISharepointTemplateProps, state: ISharepointTemplateState){
    super(props); 
     
    this.state = {

    };
  }

  public render(): React.ReactElement<ISharepointTemplateProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
        <div className="sharepointTemplate">
          <div className='ms-Grid'>
            
            <div className='ms-Grid-row'>
              <div className='ms-Grid-col'>
                <div className='welcome-text'>
                  <div className='Header-Text'>
                    <h1>Welcome Sabina</h1>
                    <p>We are excited to have you on board! To help you <br />jump-start your company journey, we included an <br /> on boarding checklist and resources for you to <br /> learn more about out team culture and leadership!</p>
                  </div>
                  <div className='Step-Image'>
                    <img src="https://boostmeup.com/blog/wp-content/uploads/2023/10/Step-by-Step.jpg" alt='Step by Step Blocks' />
                    {/* {require("../assets/Image/Step-by-Step.jpg")} */}
                  </div>
                </div>
              </div>
            </div>

              <div className='ms-Grid-col'>
                <div className='Step-Icons'> 
                  <div className='All-Icons'>
                    <Icon iconName='Globe' className='Items-Icon'/>
                    <p>Resources</p>
                  </div>

                  <div className='All-Icons'>
                    <Icon iconName='BarChart4' className='Items-Icon'/>
                    <p>Benifits</p>
                  </div>

                  <div className='All-Icons'>
                    <Icon iconName='Contact' className='Items-Icon'/>
                    <p>Directory</p>
                  </div>

                  <div className='All-Icons'>
                    <Icon iconName='OfficeStoreLogo' className='Items-Icon'/>
                    <p>Time Off</p>
                  </div>

                  <div className='All-Icons'>
                    <Icon iconName='CircleDollar' className='Items-Icon'/>
                    <p>Paystubs</p>
                  </div>

                  <div className='All-Icons'>
                    <Icon iconName='TextDocumentEdit' className='Items-Icon'/>
                    <p>Requests</p>
                  </div>
                  
                  <div className='All-Icons'>
                    <Icon iconName='Family' className='Items-Icon'/>
                    <p>Coaches</p>
                  </div>

                </div>
              </div>


          </div>
        </div>
    );
  }
}
