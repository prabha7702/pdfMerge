import * as React from "react";
import * as ReactDOM from "react-dom";
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
  type ListViewStateChangedEventArgs,
} from "@microsoft/sp-listview-extensibility";
import PDFMerge, { IPDFMergeProps } from "./components/PDFMerge/PDFMerge";

export interface IPDFMergeCommandSetProperties {}

export default class PDFMergeCommandSet extends BaseListViewCommandSet<IPDFMergeCommandSetProperties> {
  private LOG_SOURCE: string = "PDFMergeCommandSet";
  private _docConfigElement: HTMLDivElement = null;

  public async onInit(): Promise<void> {
    // console.log("PDF Merge extenion initialized")
    try {
      const command: Command = this.tryGetCommand("PDF_MERGE");
      command.visible = false;

      this.context.listView.listViewStateChangedEvent.add(
        this,
        this._onListViewStateChanged
      );

      console.info(`${this.LOG_SOURCE} (onInit) - Complete`);
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (onInit) - ${err}`);
    }
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    try {
      switch (event.itemId) {
        case "PDF_MERGE":
          this._openDialog();
          break;
        default:
          throw new Error("Unknown command");
      }
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (onExecute) - ${err}`);
    }
  }

  private _onListViewStateChanged = (
    args: ListViewStateChangedEventArgs
  ): void => {
    try {
      const command: Command = this.tryGetCommand("PDF_MERGE");
      const files = this.context.listView.selectedRows.filter(
        (row) => row.getValueByName("ContentType") === "SBS_file"
      );
      if (command) {
        command.visible =
          this.context.listView.selectedRows?.length > 1 &&
          this.context.listView.selectedRows?.length === files.length;
      }
      this.raiseOnChange();
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (_onListViewStateChanged) - ${err}`);
    }
  };

  private _openDialog(): void {
    try {
      if (this._docConfigElement == undefined) {
        this._docConfigElement = document.createElement(
          "DIV"
        ) as HTMLDivElement;
        this._docConfigElement.style.position = "relative";
        this._docConfigElement.style.display = "block";
        document.body.appendChild(this._docConfigElement);
      }

      const props: IPDFMergeProps = {
        listId: this.context.pageContext.list.id.toString(),
        documents: this.context.listView.selectedRows?.map((item) => {
          return {
            id: item.getValueByName("ID") as number,
            name: item.getValueByName("FileLeafRef"),
          };
        }),
        closePanel: this._closeConfigForm,
        context: this.context,
      };
      const element: React.ReactElement<{}> = React.createElement(
        PDFMerge,
        props
      );
      ReactDOM.render(element, this._docConfigElement);
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (_openConfigForm) - ${err}`);
    }
  }

  private _closeConfigForm = (): void => {
    if (this._docConfigElement !== undefined) {
      ReactDOM.unmountComponentAtNode(this._docConfigElement);
    }
  };
}
