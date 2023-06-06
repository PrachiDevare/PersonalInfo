import * as React from 'react';
import styles from './ManagerProfile.module.scss';
import { IManagerProfileProps } from './IManagerProfileProps';
import { MSGraphClientV3 } from "@microsoft/sp-http";


// ManagerProfile Interface
interface IProfile {
  displayName:string;
  givenName:string;
  mail:string;
  surname:string;
 

}

export default class ManagerProfile extends React.Component<IManagerProfileProps, IProfile> {

  constructor(props: IManagerProfileProps , state: IProfile) {
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
          .api("me/manager")
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
  public render(): React.ReactElement<IManagerProfileProps> {
  

    return (
      <section>
      <div>
    <h3 className ={styles.heading}> {this.props.componentTitle}</h3></div>
    <div>
            {/* <p>{this.state.displayName}</p>  */}
            <p><b>Name : </b>{this.state.givenName}&nbsp;{this.state.surname}</p>
            <p><b>Email : </b> {this.state.mail}</p> 
            </div>
     </section>
    );
  }
}
