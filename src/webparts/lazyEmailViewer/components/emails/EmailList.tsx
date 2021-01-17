import * as React from "react";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  SelectionMode,
  IColumn,
} from "office-ui-fabric-react/lib/DetailsList";
import { Fabric } from "office-ui-fabric-react/lib/Fabric";
import {
  TooltipHost,
  ITooltipHostStyles,
} from "office-ui-fabric-react/lib/Tooltip";

export interface IEmailListProps {
  emails: MicrosoftGraph.Message[];
}

const calloutProps = { gapSpace: 0 };
const hostStyles: Partial<ITooltipHostStyles> = {
  root: { display: "inline-block" },
};

export default class LazyEmailViewer extends React.Component<
  IEmailListProps,
  {}
> {
  columns: IColumn[];
  constructor(props: IEmailListProps) {
    super(props);
    this.columns = [
      {
        key: "column1",
        name: "Subject",
        minWidth: 50,
        maxWidth: 150,
        data: "string",
        isPadded: true,
        onRender: (item: MicrosoftGraph.Message) => {
          return (
            <TooltipHost
              calloutProps={calloutProps}
              styles={hostStyles}
              content={item.subject}
            >
              <span>{item.subject}</span>{" "}
            </TooltipHost>
          );
        },
      },
      {
        key: "column2",
        name: "Sender",
        minWidth: 50,
        maxWidth: 150,
        data: "string",
        isPadded: true,
        onRender: (item: MicrosoftGraph.Message) => {
          return (
            <TooltipHost
              calloutProps={calloutProps}
              styles={hostStyles}
              content={item.sender.emailAddress.address}
            >
              <span>{item.sender.emailAddress.address}</span>{" "}
            </TooltipHost>
          );
        },
      },
      {
        key: "column3",
        name: "Receive Date",
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        data: "string",
        isPadded: true,
        onRender: (item: MicrosoftGraph.Message) => {
          return (
            <TooltipHost
              calloutProps={calloutProps}
              styles={hostStyles}
              content={item.receivedDateTime}
            >
              <span>{item.receivedDateTime}</span>{" "}
            </TooltipHost>
          );
        },
      },
      {
        key: "column4",
        name: "Is read?",
        minWidth: 40,
        maxWidth: 90,
        isResizable: true,
        isCollapsible: true,
        data: "string",
        isPadded: true,
        onRender: (item: MicrosoftGraph.Message) => {
          return <span>{item.isRead ? "Yes" : "No"}</span>;
        },
      },
    ];
  }

  public render(): React.ReactElement<IEmailListProps> {
    return (
      <Fabric>
        <DetailsList
          items={this.props.emails}
          columns={this.columns}
          layoutMode={DetailsListLayoutMode.justified}
          isHeaderVisible={true}
        />
      </Fabric>
    );
  }
}
