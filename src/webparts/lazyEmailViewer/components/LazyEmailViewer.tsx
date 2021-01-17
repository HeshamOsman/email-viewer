import * as React from "react";
import styles from "./LazyEmailViewer.module.scss";
import { ILazyEmailViewerProps } from "./ILazyEmailViewerProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { MSGraphClient } from "@microsoft/sp-http";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import { List } from "office-ui-fabric-react/lib/List";
import { Environment, EnvironmentType } from "@microsoft/sp-core-library";
import EmailList from "./emails/EmailList";
import Analyser, { IAnalyzerDecisionDTO } from "./analyser/Analyser";
import { HttpClient, HttpClientResponse } from "@microsoft/sp-http";

export interface ILazyEmailViewerState {
  emails: MicrosoftGraph.Message[];
  longestEmailAnalysis: IAnalyzerDecisionDTO;
  userSendLongEmailsAnalysis: IAnalyzerDecisionDTO;
}

export default class LazyEmailViewer extends React.Component<
  ILazyEmailViewerProps,
  ILazyEmailViewerState
> {
  baseEmailAnalyserURL: string = "http://localhost:8080"; //"https://a4233b45bc5e.ngrok.io";
  constructor(props: ILazyEmailViewerProps) {
    super(props);
    this.state = {
      emails: [],
      longestEmailAnalysis: null,
      userSendLongEmailsAnalysis: null,
    };
  }
  public componentDidMount(): void {
    if (Environment.type == EnvironmentType.Local) {
      this.props.httpClient
        .get(
          "https://my.api.mockaroo.com/microsoftemail.json?key=51407670",
          HttpClient.configurations.v1
        )
        .then(
          (res: HttpClientResponse): Promise<any> => {
            return res.json();
          }
        )
        .then((response: any): void => {
          this.submitAndGetDataFromEmailAnalyser(response);
        });
    } else if (
      Environment.type == EnvironmentType.SharePoint ||
      Environment.type == EnvironmentType.ClassicSharePoint
    ) {
      this.props.mSGraphClientPromise.then((client: MSGraphClient): void => {
        client
          .api("/me/messages")
          .top(10)
          .orderby("receivedDateTime desc")
          .get((error, messages: any, rawResponse?: any) => {
            this.submitAndGetDataFromEmailAnalyser(messages.value);
          });
      });
    }
  }

  public submitAndGetDataFromEmailAnalyser(emails: MicrosoftGraph.Message[]) {
    const headers: Headers = new Headers();
    headers.append("Content-type", "application/json");
    this.props.httpClient
      .post(
        `${this.baseEmailAnalyserURL}/api/emails`,
        HttpClient.configurations.v1,
        {
          body: JSON.stringify(
            emails.map((e) => {
              return {
                id: e.id,
                emailContent: e.bodyPreview,
                sender: e.sender.emailAddress.address,
                subject: e.subject,
              };
            })
          ),
          headers,
        }
      )
      .then(
        (res: HttpClientResponse): Promise<any> => {
          return res.json();
        }
      )
      .then((response: any): void => {
        let longestEmailAnalysis: IAnalyzerDecisionDTO;
        let userSendLongEmailsAnalysis: IAnalyzerDecisionDTO;
        longestEmailAnalysis = response;

        this.props.httpClient
          .get(
            `${this.baseEmailAnalyserURL}/api/emails/analyser`,
            HttpClient.configurations.v1
          )
          .then(
            (res: HttpClientResponse): Promise<any> => {
              return res.json();
            }
          )
          .then((response: any): void => {
            userSendLongEmailsAnalysis = response;

            this.setState({
              emails,
              userSendLongEmailsAnalysis,
              longestEmailAnalysis,
            });
          });
      });
  }

  public renderError() {
    return (
      <div style={{ backgroundColor: "red" }}>
        "Can not read emails in local environment"
      </div>
    );
  }

  public render(): React.ReactElement<ILazyEmailViewerProps> {
    const {
      longestEmailAnalysis,
      emails,
      userSendLongEmailsAnalysis,
    } = this.state;

    return (
      <div className={styles.lazyEmailViewer}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>{escape(this.props.title)}</span>
              <div className={styles.subTitle}>Emails</div>
              <EmailList emails={emails} />
              <div className={styles.subTitle}>Analysis</div>
              <Analyser
                longestEmailAnalysis={longestEmailAnalysis}
                userSendLongEmailsAnalysis={userSendLongEmailsAnalysis}
              />
            </div>
          </div>
        </div>
      </div>
    );
  }
}
