import * as React from 'react';
import styles from './ReactPPPMGT.module.scss';
import { IReactPPPMGTProps } from './IReactPPPMGTProps';
import { escape } from '@microsoft/sp-lodash-subset';
import {addDays} from '@fluentui/date-time-utilities';

export default class ReactPPPMGT extends React.Component<IReactPPPMGTProps, {}> {
  public render(): React.ReactElement<IReactPPPMGTProps> {
    return (
      <div className={ styles.reactPppmgt }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
            <span className={ styles.title }>Property Pane Portal</span>
              <p className={ styles.subTitle }>Use any form control in the Property Pane.</p>
              <p className={styles.description}>MGT People Picker: {escape(this.props.mgtPeoplePicker || "")}</p>
              <p className={styles.description}>MGT Teams Channel Picker: {escape(this.props.mgtTeamsChannelPicker || "")}</p>
              <a href="https://docs.microsoft.com/en-us/graph/toolkit/overview" className={ styles.button }>
                <span className={ styles.label }>Visit the Microsoft Graph Toolkit</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
