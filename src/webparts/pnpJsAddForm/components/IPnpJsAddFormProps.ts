import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IPnpJsAddFormProps {
  description: string;
  context: WebPartContext;
  siteUrl: string;
}
