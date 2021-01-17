import { MSGraphClient,HttpClient } from '@microsoft/sp-http';

export interface ILazyEmailViewerProps {
  title: string;
  mSGraphClientPromise :Promise<MSGraphClient>;
  httpClient :HttpClient
}
