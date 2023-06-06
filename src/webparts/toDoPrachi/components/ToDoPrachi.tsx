import * as React from 'react';
import styles from './ToDoPrachi.module.scss';
import { IToDoPrachiProps } from './IToDoPrachiProps';
import { MSGraphClientV3 } from "@microsoft/sp-http";
import * as moment from 'moment';

// To do Interface
interface IToDo {
  displayName: string;
  createdDateTime: any;
  dueDateTime: any;
  title:string;
  isOwner:string;
  isShared:string;
  wellknownListName:String;
}
// All To Do Interface
interface IAllToDo {
  AllToDo: IToDo[];
}

export default class ToDoPrachi extends React.Component<IToDoPrachiProps,IAllToDo> {
  
  constructor(props: IToDoPrachiProps, state:IAllToDo) {
    super(props);
    this.state = {
      AllToDo: [],
    };
  }
  componentDidMount(): void {
    this.getMyToDo();
   
  }
  public getMyToDo() {
    this.props.context.msGraphClientFactory
      .getClient("3")
      .then((client: MSGraphClientV3): void => {
        client
          .api("me/todo/lists")
          .version("v1.0")
          .select("*")
          .get((err: any, res: any) => {
            this.setState({
              AllToDo: res.value,
            });
            console.log(this.state.AllToDo);
            // console.log(res);
            // console.log(err);
          });
      });
  }


  public render(): React.ReactElement<IToDoPrachiProps> {
   

    return (
      <section>
      <div>
   <h3 className ={styles.heading}> {this.props.componentTitle}</h3></div>
   <div className={styles["card-container"]}> 
    <div>
     {this.state.AllToDo.map((todo) => {
       return (
         <div className={styles["card-div"]}>
        
           <p><b>Display Name :</b> {todo.displayName}</p>
           <p><b>Created date time :</b> {moment(todo.createdDateTime).format("LL")}</p>
           <p><b>Due Date Time :</b> {moment(todo.dueDateTime).format("LL")}</p>
           <p><b>List Name :</b> {todo.wellknownListName}</p>
           <hr />
         </div>
       );
     })}
   </div></div></section>)
   }
  }
 
