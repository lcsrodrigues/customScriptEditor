import * as React from 'react';
import * as ReactDom from 'react-dom';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  IPropertyPaneField
} from '@microsoft/sp-property-pane';

import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'ModernScriptEditorWebpartWebPartStrings';
import ModernScriptEditorWebpart from './components/ModernScriptEditorWebpart';
import { IModernScriptEditorWebpartProps } from './components/IModernScriptEditorWebpartProps';
import DeveloperDetails from './DeveloperDetails';
import { IModernScriptEditorProps } from './IModernScriptEditorProps';

export default class ModernScriptEditorWebpartWebPart extends BaseClientSideWebPart <IModernScriptEditorProps> {
  public _propertyPaneHelper;
  private _unqiueId;

  constructor() {
    super();
    this.scriptUpdate = this.scriptUpdate.bind(this);
}

public scriptUpdate(_property: string, _oldVal: string, newVal: string) {
  this.properties.script = newVal;
  this._propertyPaneHelper.initialValue = newVal;
}

  public render(): void {
    this._unqiueId = this.context.instanceId;
    if (this.displayMode == DisplayMode.Read) {
      if (this.properties.removePadding) {
          let element = this.domElement.parentElement;
          for (let i = 0; i < 5; i++) {
              const style = window.getComputedStyle(element);
              const hasPadding = style.paddingTop !== "0px";
              if (hasPadding) {
                  element.style.paddingTop = "0px";
                  element.style.paddingBottom = "0px";
                  element.style.marginTop = "0px";
                  element.style.marginBottom = "0px";
              }
              element = element.parentElement;
          }
      }
      ReactDom.unmountComponentAtNode(this.domElement);
            this.domElement.innerHTML = this.properties.script;
            this.executeScript(this.domElement);
        } else {
            this.renderEditor();
        }
    }
    private async renderEditor() {
      const editorPopUp = await import(
          './components/ModernScriptEditorWebpart'
      );
      const element: React.ReactElement<IModernScriptEditorWebpartProps> = React.createElement(
          editorPopUp.default,
          {
              script: this.properties.script,
              title: this.properties.title,
              propPaneHandle: this.context.propertyPane,
              key: "pnp" + new Date().getTime()
          }
      );
      ReactDom.render(element, this.domElement);
       }
    
  protected get dataVersion(): Version {
      return Version.parse('1.0');
  }
  protected async loadPropertyPaneResources(): Promise<void> {
    const editorProp = await import(
        '@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor'
    );
    this._propertyPaneHelper = editorProp.PropertyFieldCodeEditor('scriptCode', {
      label: 'Edit HTML Code',
      panelTitle: 'Edit HTML Code',
      initialValue: this.properties.script,
      onPropertyChange: this.scriptUpdate,
      properties: this.properties,
      disabled: false,
      key: 'codeEditorFieldId',
      language: editorProp.PropertyFieldCodeEditorLanguages.HTML
  });
}

protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
  let webPartOptions: IPropertyPaneField<any>[] = [
      PropertyPaneTextField("title", {
          label: "Title to show in edit mode",
          value: this.properties.title
      }),
      PropertyPaneToggle("removePadding", {
          label: "Remove top/bottom padding of web part container",
          checked: this.properties.removePadding,
          onText: "Remove padding",
          offText: "Keep padding"
      }),
      PropertyPaneToggle("spPageContextInfo", {
          label: "Enable classic _spPageContextInfo",
          checked: this.properties.spPageContextInfo,
          onText: "Enabled",
          offText: "Disabled"
      }),
      this._propertyPaneHelper
  ];
  if (this.context.sdks.microsoftTeams) {
    let config = PropertyPaneToggle("teamsContext", {
        label: "Enable teams context as _teamsContexInfo",
        checked: this.properties.teamsContext,
        onText: "Enabled",
        offText: "Disabled"
    });
    webPartOptions.push(config);
}
webPartOptions.push(new DeveloperDetails());
return {
  pages: [
      {
          groups: [
              {
                  groupFields: webPartOptions
              }
          ]
      }
  ]
};
}
private evalScript(elem) {
  const data = (elem.text || elem.textContent || elem.innerHTML || "");
  const headTag = document.getElementsByTagName("head")[0] || document.documentElement;
  const scriptTag = document.createElement("script");

  for (let i = 0; i < elem.attributes.length; i++) {
      const attr = elem.attributes[i];
      if(attr.name.toLowerCase() === "onload"  ) continue; 
      scriptTag.setAttribute(attr.name, attr.value);
  }

  scriptTag.type = (scriptTag.src && scriptTag.src.length) > 0 ? "pnp" : "text/javascript";
  scriptTag.setAttribute("pnpname", this._unqiueId);

  try {
      scriptTag.appendChild(document.createTextNode(data));
  } catch (e) {
      scriptTag.text = data;
  }

  headTag.insertBefore(scriptTag, headTag.firstChild);
}

private nodeName(elem, name) {
  return elem.nodeName && elem.nodeName.toUpperCase() === name.toUpperCase();
}

private async executeScript(element: HTMLElement) {
  const headTag = document.getElementsByTagName("head")[0] || document.documentElement;
  let scriptTags = headTag.getElementsByTagName("script");
  for (let i = 0; i < scriptTags.length; i++) {
      const scriptTag = scriptTags[i];
      if(scriptTag.hasAttribute("pnpname") && scriptTag.attributes["pnpname"].value == this._unqiueId ) {
          headTag.removeChild(scriptTag);
      }            
  }

  if (this.properties.spPageContextInfo && !window["_spPageContextInfo"]) {
      window["_spPageContextInfo"] = this.context.pageContext.legacyPageContext;
  }

  if (this.properties.teamsContext && !window["_teamsContexInfo"]) {
      window["_teamsContexInfo"] = this.context.sdks.microsoftTeams.context;
  }

  (<any>window).ScriptGlobal = {};

  const scripts = [];
  const children_nodes = element.childNodes;

  for (let i = 0; children_nodes[i]; i++) {
      const child: any = children_nodes[i];
      if (this.nodeName(child, "script") &&
          (!child.type || child.type.toLowerCase() === "text/javascript")) {
          scripts.push(child);
      }
  }

  const urls = [];
  const onLoads = [];
  for (let i = 0; scripts[i]; i++) {
      const scriptTag = scripts[i];
      if (scriptTag.src && scriptTag.src.length > 0) {
          urls.push(scriptTag.src);
      }
      if (scriptTag.onload && scriptTag.onload.length > 0) {
          onLoads.push(scriptTag.onload);
      }
  }

  let oldamd = null;
  if (window["define"] && window["define"].amd) {
      oldamd = window["define"].amd;
      window["define"].amd = null;
  }

  for (let i = 0; i < urls.length; i++) {
      try {
          let scriptUrl = urls[i];
          const prefix = scriptUrl.indexOf('?') === -1 ? '?' : '&';
          scriptUrl += prefix + 'pnp=' + new Date().getTime();
          await SPComponentLoader.loadScript(scriptUrl, { globalExportsName: "ScriptGlobal" });
      } catch (error) {
          if (console.error) {
              console.error(error);
          }
      }
  }
  if (oldamd) {
      window["define"].amd = oldamd;
  }

  for (let i = 0; scripts[i]; i++) {
      const scriptTag = scripts[i];
      if (scriptTag.parentNode) { scriptTag.parentNode.removeChild(scriptTag); }
      this.evalScript(scripts[i]);
  }
  for (let i = 0; onLoads[i]; i++) {
      onLoads[i]();
  }
}
}