import * as React from "react";
import { IHelloWorldProps } from "./IHelloWorldProps";
import styles from "./HelloWorld.module.scss";

export default class HelloWorld extends React.Component<IHelloWorldProps, {}> {

  public render(): React.ReactElement<IHelloWorldProps> {
    const { test2, test3, _environmentMessage, userDisplayName, test1, description, hasTeamsContext, isDarkTheme, } = this.props;

    return (
      <div>
        <section className={`${styles.helloWorld} ${hasTeamsContext ? styles.teams : ''}`}>
          <div className={styles.welcome}>
            <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
            <h2>Well done, {userDisplayName}!</h2>
            <div>{_environmentMessage}</div>
          </div>
          <div>
            <h3>Welcome to SharePoint Framework!</h3>
            <div>Web part description: <strong>{description}</strong></div>
            <div>Web part test: <strong>{test1 ? 'Checked' : 'Unchecked'}</strong></div>
            <div>Dropdown selection: <strong>{test2}</strong></div>
            <div>Toggle status: <strong>{test3 ? 'Enabled' : 'Disabled'}</strong></div>
            <div>Loading from: <strong>{this.props.pageContext.web.title}</strong></div>
          </div>
          <div id="spListContainer" />
        </section>
      </div>
    );
  }
}
