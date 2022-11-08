import { IPropertyPaneAccessor } from "@microsoft/sp-webpart-base";
export interface IModernScriptEditorWebpartProps {
  script: string;
  title: string;
  propPaneHandle: IPropertyPaneAccessor;  
}