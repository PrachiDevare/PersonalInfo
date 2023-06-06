import * as React from 'react';
import styles from './MyAccount.module.scss';
import { IMyAccountProps } from './IMyAccountProps';
import { MSGraphClientV3 } from "@microsoft/sp-http";

// Profile Interface
interface IProfile {
  displayName:string;
  givenName:string;
  mail:string;
  surname:string;
 

}
// // All Items Interface
// interface IAllProfile {
//   AllProfile: IProfile[];
// }

export default class MyAccount extends React.Component<IMyAccountProps, IProfile> {

  constructor(props: IMyAccountProps, state: IProfile) {
    super(props);
    this.state = 
      {displayName:"",mail:"",givenName:"",surname:""}
    
  }
  componentDidMount(): void {
    this.getMyProfile();
   
  }
  public getMyProfile() {
    this.props.context.msGraphClientFactory
      .getClient("3")
      .then((client: MSGraphClientV3): void => {
        client
          .api("/me")
          .version("v1.0")
          .select("dispayName,givenName,mail,surname,userPrincipalName")
          .get((err: any, res: any) => {
            this.setState({
              displayName:res.displayName,
              mail:res.mail,
              givenName:res.givenName,
              surname:res.surname
            })
            console.log(res);
           
            // console.log(err);
          });
      });
  }
  public render(): React.ReactElement<IMyAccountProps> {
   
       return (
       <section>
        <div>
      <h3 className ={styles.heading}> {this.props.componentTitle}</h3></div>
      <div>
              {/* <p>{this.state.displayName}</p>  */}
              <p><b>Name : </b>{this.state.givenName}&nbsp;{this.state.surname}</p>
              <p><b></b> {this.state.mail}</p> 
              </div>
       </section>
    );
  }
}
