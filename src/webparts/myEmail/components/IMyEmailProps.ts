import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IMyEmailProps {
  componentTitle: string;
  context:WebPartContext;
  mail:Boolean;
}
