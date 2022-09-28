import * as React from 'react';
import styles from './Grid.module.scss';
import { IGridProps } from './IGridProps';
import { useEffect, useState } from "react"
import { sortBy } from '@microsoft/sp-lodash-subset';

import { IIconProps } from 'office-ui-fabric-react';
import { IconButton } from 'office-ui-fabric-react/lib/Button';
import { Spinner } from 'office-ui-fabric-react';
import { Label } from 'office-ui-fabric-react';

const collapseIcon: IIconProps = { iconName: 'CollapseContentSingle' };
const exploreIcon: IIconProps = { iconName: 'ExploreContentSingle' };

const colors: string[] = ["#E9EBF5", "#CFD5EA"];

function App(props: IGridProps) {
  const [data, setData] = useState([]);
  const [collapsed, setCollapsed] = useState(props.collapsed);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    const fetchData = async () => {
      try {
        const items = await props.service.getItems(props.list, props.fields, props.fields.filter(x => !!x.groupByField).pop()?.groupByField, props.orderBy, props.fullWidth);
        setData(items);
      } catch (error) {
          console.log("error", error);
      }

      setLoading(false);
    };
    fetchData();
  }, []);

  const fields = sortBy(props.fields, x => x.sortIdx);

  let currentRow = 1;

  return (
    <div className={`${styles.grid} ${props.hasTeamsContext ? styles.teams : ''}`}>

      <div className={styles.webPartHeaderContainer}>
        <div className={`${styles.webPartHeader}`}>
            <div className={styles.webPartTitle}>
              <span role="heading" aria-level={2}>{props.title}</span>
              <IconButton
                iconProps={collapsed ? exploreIcon : collapseIcon}
                ariaLabel="Collapse / Explore"
                onClick={() => { setCollapsed(!collapsed);}}
              />
          </div>
        </div>
      </div>

      {!collapsed && (
        <div className={styles.header}>
          {fields.map(f => (
            <div className={styles['header-column']} style={{width: f.width + '%'}}>{f.title}</div>
          ))}
        </div>
      )}

      {!collapsed && loading && (
        <div className={styles.row}>
          <div className={styles['column']} style={{width: '100%'}}>
            <Label>Loading...</Label>
            <Spinner label="I am definitely loading..." />
          </div>
        </div>
      )}

      {!collapsed && !loading && data.map(group => {
        const groupByField = fields.filter(x => !!x.groupByField).pop();

        if (!groupByField)
        {
          return (
            group.map(item => {
              return (
                <div className={styles.row} style={{backgroundColor: colors[currentRow++ % 2]}}>
                  {!!props.fullWidth && item[props.fullWidth] && (
                    <div className={styles['column']} style={{width: '100%'}} dangerouslySetInnerHTML={ {__html: item[fields[0].field]}}></div>
                  )}
                  {(!props.fullWidth || !item[props.fullWidth]) && (
                    <>
                    {fields.map(f => (
                      <div className={styles['column']} style={{width: f.width + '%'}} dangerouslySetInnerHTML={ {__html: item[f.field]}}></div>
                    ))}
                    </>
                  )}
                </div>
              )
            })
          )
        }

        const spareWidth = 100 - groupByField.width;

        return (
          <div className={styles.row} style={{backgroundColor: colors[currentRow++ % 2]}}>

            {!!props.fullWidth && group[0][props.fullWidth] && (
                <div className={styles['column']} style={{width: '100%'}} dangerouslySetInnerHTML={ {__html: group[0][groupByField.field]}}></div>
            )}

            {(!props.fullWidth || !group[0][props.fullWidth]) && (
            <>
              <div className={styles['columnGroupName']} style={{width: groupByField.width + '%'}} ><div dangerouslySetInnerHTML={ {__html: group[0][groupByField.field]}}></div></div>
              <div className={styles['columnGroup']} style={{width: 100 - groupByField.width + '%'}}>
              {group.map((item, index: number) => {
                if (index === 0){currentRow--;}
                return (
                  <div className={styles.row} style={{backgroundColor: colors[currentRow++ % 2], flex: "1 1", flexBasis: "content"}}>
                    {fields.filter(x => x.field !== groupByField.field).map(f => (
                      <div className={styles['columnInner']} style={{width:  f.width*100/spareWidth + '%'}} dangerouslySetInnerHTML={ {__html: item[f.field]}}></div>
                    ))}
                  </div>
                )
              })}
              </div>
            </>
            )}

          </div>)
      })}

      {!collapsed && props.footer && (
      <div className={styles.row} style={{backgroundColor: colors[currentRow++ % 2]}}>
         <div className={styles['column']} style={{width: '100%'}} dangerouslySetInnerHTML={ {__html: props.footer }}></div>
      </div>
      )}

    </div>
  );
}

export default App;