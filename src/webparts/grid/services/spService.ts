import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { PageContext } from "@microsoft/sp-page-context";

import { spfi, SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import { ISelectedField } from '../GridWebPart';
import { groupBy, sortBy } from '@microsoft/sp-lodash-subset';

export interface ISharePointService {
    getItems: (list: string, fields: any[], groupByField: string, orderBy: string, fullWidth: string) => Promise<any>;
}

export class SPService implements ISharePointService {
    public static readonly serviceKey: ServiceKey<ISharePointService> = ServiceKey.create<ISharePointService>('SPFx:SharePointService', SPService);
    private _sp: SPFI;
  
    constructor(serviceScope: ServiceScope) {
      serviceScope.whenFinished(() => {
        const pageContext = serviceScope.consume(PageContext.serviceKey);
        this._sp = spfi().using(SPFx({ pageContext }));
      });
    }

    public async getItems(listId: string, fields: ISelectedField[], groupByField: string, orderBy: string, fullWidth: string): Promise<any>
    {
        let selectFields = [...fields.map(x => x.field)];

        if (groupByField) {
            selectFields.push(groupByField);
        }

        if (orderBy){
            selectFields.push(orderBy);
        }

        if (fullWidth){
            selectFields.push(fullWidth);
        }

        const items = await this._sp.web.lists.getById(listId).items.select(...selectFields).orderBy(orderBy)();
        const groupedList = groupBy(items, groupByField);
        const sorted = sortBy(groupedList, group => items.indexOf(group[0])); // maintain order

        return sorted;
    }
}