import * as React from 'react';
import styles from './MsGraphApiSample.module.scss';
import { IMsGraphApiSampleProps } from './IMsGraphApiSampleProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { graph } from "@pnp/graph";
import "@pnp/graph/users";
import "@pnp/graph/insights";
import "@pnp/graph/batch";
import "@pnp/graph/messages";
import { IUserItem } from './IUserItem';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';

export interface IMsGraphApiSampleState {
  loading: boolean;
  result: string;
  users: Array<IUserItem>;
  displayName: string;
 
}



export default class MsGraphApiSample extends React.Component<IMsGraphApiSampleProps, IMsGraphApiSampleState> {

  constructor(props: IMsGraphApiSampleProps) {
    super(props);
    this.state = {
        loading: false,
        result:'',
        users: [],
        displayName:'',
        
    };
  
}


private _getCurrentUserInfo = async () => {
  this.setState({
      loading: true,
      result: '',
      displayName:''
  });
  let userInfo: any = await this.props.client.api('/users/anthony@asmurph.onmicrosoft.com').version('/beta').get();
  this.setState({
      loading: false,
      displayName:'',
      result: JSON.stringify(userInfo, undefined, 2)
  });
}

private _getPnPUserInfo = async () => {
  this.setState({
      loading: true,
      result: ''
  });
  let userInfo: any = await graph.me.get();
  this.setState({
      loading: false,
      result: JSON.stringify(userInfo, undefined, 2)
  });
}
private _getBatchResponse = async () => {
  this.setState({
      loading: true,
      result: ''
  });
  let batchReqests: any = {
      "requests": [
          {
              "url": "/me?$select=displayName,jobTitle,userPrincipalName",
              "method": "GET",
              "id": "1"
          },
          {
              "url": "/me/messages?$filter=importance eq 'high'&$select=from,subject",
              "method": "GET",
              "id": "2",
              "DependsOn": [
                  "1"
              ]
          },
          {
              "url": "/me/events?$select=subject,organizer",
              "method": "GET",
              "id": "3",
              "DependsOn": [
                  "2"
              ]
          }
      ]
  };
  let batchResponse: any = await this.props.client.api('$batch').post(batchReqests);
  this.setState({
      loading: false,
      result: JSON.stringify(batchResponse, undefined, 2)
      
  });
}


  public render(): React.ReactElement<IMsGraphApiSampleProps> {
    const { items, loading, result } = this.state;
    return (
      <div className={ styles.msGraphApiSample }>
        <div className={ styles.container }>
        <div className={styles.row}>
                  <div className={styles.column}>
                      <div>
                          <div><h3>Using MSGraph Client</h3></div>
                      </div>
                      <DefaultButton onClick={this._getCurrentUserInfo}>Get User Info</DefaultButton>
         
                      <DefaultButton onClick={this._getBatchResponse}>Get Batch Response</DefaultButton>
                      <div>
                          <div><h3>Using PnPGraph</h3></div>
                      </div>
                      <DefaultButton onClick={this._getPnPUserInfo}>Get User Info</DefaultButton>
                     
                      {this.state.loading &&
                          <div><h4>Please wait, loading...</h4></div>
                      }
                      {this.state.result &&
                          <div style={{ wordBreak: 'break-word', maxHeight: '400px', overflowY: 'auto' }}>
                              <pre>{this.state.result}</pre>
                              <pre>{this.state.displayName}</pre>
                          </div>
                      }
                       <ul>
          {items.map(item => (
            <li key={item.id}>
              {item.name} {item.price}
            </li>
          ))}
        </ul>
                  </div>
              </div>
        </div>
      </div>
    );
  }
}
