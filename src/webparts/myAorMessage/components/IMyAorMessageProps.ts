import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';

export interface IMyAorMessageProps {
  description: string;
  messagingClient: AadHttpClient;
}
