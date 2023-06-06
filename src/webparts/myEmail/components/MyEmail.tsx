import * as React from 'react';
import styles from './MyEmail.module.scss';
import { IMyEmailProps } from './IMyEmailProps';
import * as moment from "moment";
import { MSGraphClientV3 } from "@microsoft/sp-http";

// Email Interface
interface IEmails {
  subject: string;
  webLink: string;
  from: {
    emailAddress: {
      name: string;
      address: string;
    };
  };
  receivedDateTime: any;
  bodyPreview: string;
  isRead: any;
}
// All Items Interface
interface IAllItems {
  AllEmails: IEmails[];
}
export default class MyEmail extends React.Component<IMyEmailProps, IAllItems> {

  constructor(props: IMyEmailProps, state: IAllItems) {
    super(props);
    this.state = {
      AllEmails: [],
    };
  }
  componentDidMount(): void {
    this.getMyEmails();
   
  }
  
  public getMyEmails() {
    this.props.context.msGraphClientFactory
      .getClient("3")
      .then((client: MSGraphClientV3): void => {
        client
          .api("/me/messages")
          .version("v1.0")
          .select("subject,webLink,from,receivedDateTime,isRead,bodyPreview")
          .get((err: any, res: any) => {
            this.setState({
              AllEmails: res.value,
            });
            console.log(this.state.AllEmails);
            // console.log(res);
            // console.log(err);
          });
      });
  }
  public render(): React.ReactElement<IMyEmailProps> {
    
    return (<section>
      {/* <div>
      <h3 className ={styles.heading}> {this.props.componentTitle}</h3></div>
        <div>
      {this.state.AllEmails.map((email) => {
        return ( 
         <div>
          
          {email.isRead?(
             <div className={styles["card-div"]}>
              <img src='https://techsavvydotlife.files.wordpress.com/2022/04/mail-app-header.jpg'width={50}></img>
              

              <h3>{email.from==undefined?"":email.from.emailAddress.name}</h3>
             <p>{email.subject}</p>
            <p>{moment(email.receivedDateTime).format("LL")}</p>
            <p>{email.bodyPreview}</p>
            <button
              onClick={() => {
                window.open(email.webLink, "_blank");
              }}
            >
              {" "}
              Open email in new tab
            </button>
            
            <hr />
          </div>
          ):
          (<div className={styles["card-div"]}>
            <img src = 'https://img.buzzfeed.com/buzzfeed-static/static/2016-07/8/13/campaign_images/buzzfeed-prod-fastlane02/heres-how-i-can-tell-if-someone-read-my-email-2-7107-1468000476-0_dblbig.jpg?resize=1200:*' width={50}></img>
            <h3>{email.from==undefined?"":email.from.emailAddress.name}</h3>
            <p>{email.subject}</p>
            <p>{moment(email.receivedDateTime).format("LL")}</p>
            <p>{email.bodyPreview}</p>
            <button
              onClick={() => {
                window.open(email.webLink, "_blank");
              }}
            >
              {" "}
              Open email in new tab
            </button>  
            <hr />
          </div>)}
        
         </div>
        );
      })}
    </div> */}


{/* circleShape */}
<div>
      <h3 className ={styles.heading}> {this.props.componentTitle}</h3></div>
       
<div>
        {this.state.AllEmails.map((email) => {
          return (
            <div className={styles["card-div"]}
            style ={{
              backgroundColor:email.isRead == false?"rgb(226, 226, 226)":"rgb(255,255,255)",
             }}>
               <p className={styles.circleShape}
           
           style={{backgroundColor:email.isRead == false? "red":"green",
          }}></p>

           <div>
            <p>{email.from==undefined?"":email.from.emailAddress.name}</p>
              <p>{email.subject}</p>
              <p>{moment(email.receivedDateTime).format("LL")}</p>
              <p>{email.bodyPreview}</p>
              <button
                onClick={() => {
                  window.open(email.webLink, "_blank");
                }}
              >
                {" "}
                open email in new tab
              </button>
              <hr />
            </div></div>
          );
        })}
      </div>
    
    </section>
    );
  }
}
