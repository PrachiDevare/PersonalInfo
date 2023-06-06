import * as React from 'react';
import styles from './MyPlanner.module.scss';
import { IMyPlannerProps } from './IMyPlannerProps';
import { MSGraphClientV3 } from "@microsoft/sp-http";
import * as moment from "moment";

// Planner Interface
interface IPlanner {
  title: string;
  createdDateTime: any;
  priority:number;
  dueDateTime: any;
  previewType: string;
  
  
}
// All planner Interface
interface IAllPlanner {
  AllPlanner: IPlanner[];
}

export default class MyPlanner extends React.Component<IMyPlannerProps, IAllPlanner> {
  
  constructor(props: IMyPlannerProps, state: IAllPlanner) {
    super(props);
    this.state = {
      AllPlanner: [],
    };
  }
  componentDidMount(): void {
    this.getMyPlanner();
   
  }
  public getMyPlanner() {
    this.props.context.msGraphClientFactory
      .getClient("3")
      .then((client: MSGraphClientV3): void => {
        client
          .api("me/planner/tasks")
          .version("v1.0")
          .select("title,createdDateTime,priority,dueDateTime,previewType,assignments")
          .get((err: any, res: any) => {
            this.setState({
              AllPlanner: res.value,
            });
            console.log(this.state.AllPlanner);
            // console.log(res);
            // console.log(err);
          });
      });
  }
  public render(): React.ReactElement<IMyPlannerProps> {
   

    return (
     <section>
         <div>
      <h3 className ={styles.heading}> {this.props.componentTitle}</h3></div>
      <div className={styles["card-container"]}> 
       <div>
        {this.state.AllPlanner.map((profile) => {
          return (
            <div className={styles["card-div"]}>
              <p><b>Title :</b> {profile.title}</p>
              <p><b>Created date time :</b> {moment(profile.createdDateTime).format("LL")}</p>
              <p><b>Due Date Time :</b> {moment(profile.dueDateTime).format("LL")}</p>
              <p><b>Priority :</b> {profile.priority}</p>
              <p><b>Preview type :</b> {profile.previewType}</p>
              <hr />
            </div>
          );
        })}
      </div></div>





      
     </section>
     
    
  );
  }
}
