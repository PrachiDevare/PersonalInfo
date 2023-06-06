import * as React from 'react';
import styles from './MyDailyTask.module.scss';
import { IMyDailyTaskProps } from './IMyDailyTaskProps';
import { MSGraphClientV3 } from "@microsoft/sp-http";
import * as moment from "moment";

// Daily Task Interface
interface IDailyTask {
subject: string;

start: {
    dateTime: any,
    timeZone: any
  };
end: {
    dateTime: any,
    timeZone: any
};
organizer: {
  emailAddress: {
      name: string,
      address: string
  }
}
}
// All dailyTask Interface
interface IAllDailyTask {
  AllDailyTask: IDailyTask[];
}


export default class MyDailyTask extends React.Component<IMyDailyTaskProps, IAllDailyTask> {
 
  constructor(props: IMyDailyTaskProps, state: IAllDailyTask) {
    super(props);
    this.state = {
      AllDailyTask: [],
    };
  }
  componentDidMount(): void {
    this.getDailyTask();
   
  }
  public getDailyTask() {
    this.props.context.msGraphClientFactory
      .getClient("3")
      .then((client: MSGraphClientV3): void => {
        client
          .api("/me/events?$select=subject,body,bodyPreview,organizer,attendees,start,end,location")
          .version("v1.0")
          .select("subject,start,end,organizer")
          .get((err: any, res: any) => {
            this.setState({
              AllDailyTask: res.value,
            });
            console.log(this.state.AllDailyTask);
            // console.log(res);
            // console.log(err);
          });
      });
  }

  public render(): React.ReactElement<IMyDailyTaskProps> {
    

    return (
      <div><h3 className ={styles.heading}> {this.props.componentTitle}</h3>
      <div>
      {this.state.AllDailyTask.map((DailyTask) => {
        return (
          <div className={styles["card-div"]}>
            <p><b>Task Name : {DailyTask.subject}</b></p>
            <p><b>Start Time : </b>{moment(DailyTask.start.dateTime).format("LLL")}&nbsp;<b>{DailyTask.start.timeZone}</b></p>
            <p><b>End Time : </b>{moment(DailyTask.end.dateTime).format("LLL")}&nbsp;<b>{DailyTask.end.timeZone}</b></p>
            <p><b>Organizer Name : </b>{DailyTask.organizer.emailAddress.name}</p>
            <p><b>Email Address : </b>{DailyTask.organizer.emailAddress.address}</p>
           
            <hr />
          </div>
        );
      })}
    </div></div>
    );
  }
}
