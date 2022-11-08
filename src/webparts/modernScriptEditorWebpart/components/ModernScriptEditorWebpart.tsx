import * as React from 'react';
import styles from './ModernScriptEditorWebpart.module.scss';
import { IModernScriptEditorWebpartProps } from './IModernScriptEditorWebpartProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";


export default class ModernScriptEditorWebpart extends React.Component<IModernScriptEditorWebpartProps,any> {

  constructor(props: IModernScriptEditorWebpartProps, state: any) {
    super(props);
    this._showDialog = this._showDialog.bind(this);
        this.state = {};
  }
  public componentDidMount(): void {
    this.setState({ script: this.props.script, loaded: this.props.script });
}

private _showDialog() {
    this.props.propPaneHandle.open();
}

  public render(): React.ReactElement<IModernScriptEditorWebpartProps> {
    const viewMode = <span dangerouslySetInnerHTML={{ __html: this.state.script }}></span>;

    return (
      <div className='ms-Fabric'>
                <Placeholder iconName='JS'
                    iconText={this.props.title}
                    description='Please configure the web part'
                    buttonLabel='Edit markup'
                    onConfigure={this._showDialog} />
                {viewMode}
            </div>);    
  }
}