import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from '@microsoft/sp-http';  

export interface INyRpgAdminSearchProps {
  context: WebPartContext;  
  spHttpClient: SPHttpClient;
  AdminGroupID: number;
}
