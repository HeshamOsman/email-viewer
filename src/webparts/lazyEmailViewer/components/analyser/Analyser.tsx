import * as React from "react";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import {
  DocumentCard,
  DocumentCardActivity,
  DocumentCardDetails,
  DocumentCardPreview,
  DocumentCardTitle,
  IDocumentCardPreviewProps,
  DocumentCardType,
  IDocumentCardActivityPerson,
} from "office-ui-fabric-react/lib/DocumentCard";
import { Stack, IStackTokens } from "office-ui-fabric-react/lib/Stack";
import { getTheme } from "office-ui-fabric-react//lib/Styling";

export interface IAnalyzerDecisionDTO {
  analysisMessage: string;
  analysisRecommendedAction: string;
}

export interface IAnalysertProps {
  longestEmailAnalysis: IAnalyzerDecisionDTO;
  userSendLongEmailsAnalysis: IAnalyzerDecisionDTO;
}

const stackTokens: IStackTokens = { childrenGap: 0 };
const theme = getTheme();
const { palette, fonts } = theme;

const ano: IDocumentCardActivityPerson = {
  name: "Anonymous Anonymous",
  profileImageSrc: "",
  initials: "NA",
};

export default class Analyser extends React.Component<IAnalysertProps, {}> {
  public render(): React.ReactElement<IAnalysertProps> {
    const { longestEmailAnalysis, userSendLongEmailsAnalysis } = this.props;
    if (longestEmailAnalysis == null && userSendLongEmailsAnalysis == null) {
      return <div></div>;
    } else {
      return (
        <Stack tokens={stackTokens}>
          <DocumentCard type={DocumentCardType.compact}>
            <DocumentCardDetails>
              <DocumentCardTitle title={longestEmailAnalysis.analysisMessage} />
              <DocumentCardActivity
                activity={
                  "Recommanded Action: " +
                  longestEmailAnalysis.analysisRecommendedAction
                }
                people={[ano]}
              />
            </DocumentCardDetails>
          </DocumentCard>
          <DocumentCard type={DocumentCardType.compact}>
            <DocumentCardDetails>
              <DocumentCardTitle
                title={userSendLongEmailsAnalysis.analysisMessage}
              />
              <DocumentCardActivity
                activity={
                  "Recommanded Action: " +
                  userSendLongEmailsAnalysis.analysisRecommendedAction
                }
                people={[ano]}
              />
            </DocumentCardDetails>
          </DocumentCard>
        </Stack>
      );
    }
  }
}
