import * as React from 'react';
import styles from './UrgentMessage.module.scss';
import { IUrgentMessageProps } from './IUrgentMessageProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { FunctionComponent, useEffect, useState } from 'react';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Checkbox, ITheme } from 'office-ui-fabric-react';
import { IReadonlyTheme } from "@microsoft/sp-component-base";

const UrgentMessage: FunctionComponent<IUrgentMessageProps> = (props) => {

  const [showMessage, setShowMessage] = useState<boolean>(false);

  const { semanticColors }: IReadonlyTheme = props.themeVariant;

  useEffect(() => {

    console.log('About to fetch data');
    fetchData();
 
  }, []);

  const fetchData = async () => {
    const items: any[] = await sp.web.lists.getById(props.list).items.top(1).filter(`Title eq '${props.currentUser}'`).get();

    if (items.length == 0) {
      setShowMessage(true);
    }

    console.log('Items', items)
  }

  function _onChange(ev: React.FormEvent<HTMLElement>, isChecked: boolean) {
    console.log(`The option has been changed to ${isChecked}.`);

    sp.web.lists.getById(props.list).items.add({
      "Title": props.currentUser
    })

    setShowMessage(false); 

  }

  return (
    <div className={styles.urgentMessage} style={{ backgroundColor: semanticColors.bodyBackground }}>
      
          {showMessage &&
              <div style= { { color: semanticColors.bodyText } }>
              <span>{props.message}</span>
              <Checkbox theme={(props.themeVariant as ITheme)} label={props.label} onChange={_onChange} />
              </div>

          }
     
    </div>
  );
}
 
export default UrgentMessage;


