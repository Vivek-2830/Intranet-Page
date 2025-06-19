import * as React from 'react';
import styles from './OnboardingResources.module.scss';
import { IOnboardingResourcesProps } from './IOnboardingResourcesProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';
import * as moment from 'moment';
import { Icon } from 'office-ui-fabric-react';

export interface IOnboardingResourcesState {

}

require("../assets/css/style.css");

export default class OnboardingResources extends React.Component<IOnboardingResourcesProps, IOnboardingResourcesState> {

  constructor(props: IOnboardingResourcesProps, state: IOnboardingResourcesState) {
    super(props);

    this.state = {

    };
  }

  public render(): React.ReactElement<IOnboardingResourcesProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
        <div className="onboardingResources">
          <div className='ms-Grid'>
            <div className='ms-Grid-row'>
              <div className='ms-Grid-col'>
                <div className='Onboarding'>
                  <h2>ONBOARDING RESOURCES</h2>
                    <div className='Onboarding-Text'> 
                      <div className='Resources-Items'>
                        <img src="https://www.nestleprofessional.com.au/sites/default/files/2022-05/nestle-brands-beverages.jpg" alt="Meet New Hires"/>
                        <p>MEET NEW HIRES</p>
                      </div>

                      <div className='Resources-Items'>
                        <img src="https://trajanwealth.com/wp-content/uploads/2023/08/series-of-doors.jpg" alt="Meet New Hires"/>
                        <p>CAREERS</p>
                      </div>

                      <div className='Resources-Items'>
                        <img src="https://minoritynurse.com/wp-content/uploads/2020/09/Online-Image-1.jpg" alt="Meet New Hires"/>
                        <p>LEARNING</p>
                      </div>
                    </div>
                </div>

                <footer className="footer">
                    <div className='footer-text'>Copyright (c) 2024 ORIGAMI</div>
                    <div className='contact-text'>Contact: vivekg1082004@outlook.com</div>
                </footer>

              </div>
            </div>
          </div>
        </div>
    );
  }
}
