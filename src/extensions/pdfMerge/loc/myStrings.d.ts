declare interface IPDFMergeCommandSetStrings {
  PanelHeader: string;
  DragDropInfo: string;
  Note: string;
  PasswordProtectionInfo: string;
  NewFileName: string;
  PDFExtension: string;
  DeleteInfo: string;
  PreviewInfo: string;
  MergeInfo: string;
  MergeandDeleteInfo: string;
  Merge:string;
  Cancel:string;
  LoadingInfo: string;
  ErrorInfo: string
}

declare module "PDFMergeCommandSetStrings" {
  const strings: IPDFMergeCommandSetStrings;
  export = strings;
}
