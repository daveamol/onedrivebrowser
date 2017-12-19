import * as React from 'react';
import './App.css';
import StatusBar from '../StatusBar/StatusBar';
import DriveBrowser from '../DriveBrowser/DriveBrowser';
import { User } from 'msal/lib-commonjs/User';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { getUserAgentApplication, applicationConfig } from '../../utils/AuthUtils';
import { initializeIcons } from '@uifabric/icons';
import * as Msal from 'msal';
import * as MsGraph from '@microsoft/microsoft-graph-client';
// import { Icon } from 'office-ui-fabric-react/lib/Icon';

// Talk About libs used
// Talk about autobind
// Talk about state interfaces having null lot of the times
// Talk about how a new adal v1 app is created to deploy your app to cloud

interface AppState {
  user: User | null;
  token: string;
  client: MsGraph.Client | null;
}

class App extends React.Component<any, AppState> {
  private userAgentApplication: Msal.UserAgentApplication;

  constructor(props: any) {
    super(props);
    this.getGraphToken();
    initializeIcons();
    this.state = { user: this.userAgentApplication.getUser(), token: '', client: null };
  }

  render() {
    let button;
    let driveBrowser;
    if (!this.state.user) {
      button = (
        <DefaultButton
          primary={true}
          text="Sign In"
          onClick={this.logIn}
        />
      );
      driveBrowser = null;
    } else {
      button = (
        <DefaultButton
          primary={true}
          text="Logout"
          onClick={this.logOut}
          className="logout-button"
        />
      );
      driveBrowser = <DriveBrowser client={this.state.client} />;
    }

    return (
      <div className="App">
        <Fabric>
          {/* <Icon iconName="onedrive" className="ms-BrandIcon--icon96" /> */}
          <StatusBar />
          {button}
          {driveBrowser}
        </Fabric>
      </div>
    );
  }

  @autobind
  public authCallback(errorDesc: string, authToken: string, error?: string, tokenType?: string) {
    if (authToken) {
      console.log('Callback auth token: ' + authToken);
      this.getGraphToken();
    } else {
      console.log(error + ':' + errorDesc);
      this.setState({ user: null });
    }
  }

  @autobind
  public getGraphToken() {
    this.userAgentApplication = getUserAgentApplication(this.authCallback);
    this.userAgentApplication.acquireTokenSilent(applicationConfig.graphScopes)
      .then(graphToken => {
        this.setGraphTokenAndClient(graphToken);
      }, error => {
        this.userAgentApplication.acquireTokenRedirect(applicationConfig.graphScopes);
      });
  }

  @autobind
  public setGraphTokenAndClient(graphToken: string) {
    let graphClient: MsGraph.Client = MsGraph.Client.init({
      authProvider: (done) => {
        done(null, graphToken);
      }
    });
    console.log('MS Graph token: ' + graphToken);
    this.setState({
      token: graphToken, client: graphClient
    });
  }

  @autobind
  private logIn() {
    this.userAgentApplication.loginRedirect(applicationConfig.graphScopes);
  }

  @autobind
  private logOut() {
    this.userAgentApplication.logout();
  }
}

export default App;
