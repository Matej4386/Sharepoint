import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';

export interface IConfigureWebPartProps {
  webPartContext: WebPartContext;
  title: string;
  description?: string;
  buttonText?: string;
}

const ConfigureWebPart: React.SFC<IConfigureWebPartProps> = (props) => {
  const {
    webPartContext,
    title,
    description,
    buttonText
  } = props;
  return (
    <div >
      <div >{title}</div>
      <div >
        <MessageBar messageBarType={MessageBarType.info} >
          {description ? description : 'Please configure this web part\'s properties first.'}
        </MessageBar>
      </div>
      <div >
        <PrimaryButton
          iconProps={{ iconName: 'Edit' }}
          onClick={(e) => { e.preventDefault(); webPartContext.propertyPane.open(); }}
        >
          {buttonText ? buttonText : 'Configure Web Part'}
        </PrimaryButton>
      </div>
    </div>
  );
};

export default ConfigureWebPart;
