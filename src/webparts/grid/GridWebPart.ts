import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'GridWebPartStrings';
import Grid from './components/Grid';
import { IGridProps } from './components/IGridProps';

import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { FieldPicker, IFieldPickerProps } from "@pnp/spfx-controls-react/lib/FieldPicker";
import { FieldsOrderBy } from '@pnp/spfx-controls-react/lib/services/ISPService';
import { ISPField } from '@pnp/spfx-controls-react';
import { IColumnReturnProperty, PropertyFieldColumnPicker, PropertyFieldColumnPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldColumnPicker';
import { PropertyFieldCodeEditor, PropertyFieldCodeEditorLanguages } from '@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor';
import { ISharePointService, SPService } from './services/spService';

export interface IGridWebPartProps {
  description: string;
  list: string; // Stores the list ID(s)
  fields: ISelectedField[];
  fullWidth: string;
  orderBy: string;
  footer: string;
  collapsed: boolean;
}

export interface ISelectedField {
  title: string;
  field: string;
  groupByField: string;
  width: number;
  sortIdx: number;
}

export default class GridWebPart extends BaseClientSideWebPart<IGridWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _service: ISharePointService;


  public render(): void {
    const element: React.ReactElement<IGridProps> = React.createElement(
      Grid,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        list: this.properties.list,
        fields: this.properties.fields,
        fullWidth: this.properties.fullWidth,
        orderBy: this.properties.orderBy,
        service: this._service,
        title: this.properties.description,
        footer: this.properties.footer,
        collapsed: this.properties.collapsed
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();
    this._service = new SPService(this.context.serviceScope);

    return super.onInit();
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Grid Configuration"
          },
          groups: [
            {

              //add group by, add order by
              groupName: "Settings",
              groupFields: [
                PropertyPaneTextField('description', {
                  label: "Title"
                }),
                PropertyPaneToggle("collapsed", {
                  label: "By default collaped.",
                }),
                PropertyFieldListPicker('list', {
                  label: 'Select a list',
                  selectedList: this.properties.list,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,  
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId'
                }),
                PropertyFieldCollectionData("fields", {
                  key: "fields",
                  label: "Fields data",
                  panelHeader: "Fields data panel header",
                  manageBtnLabel: "Manage fields data",
                  value: this.properties.fields,
                  enableSorting: true,
                  fields: [
                    {
                      id: "title",
                      title: "Title",
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: "field",
                      title: "Field",
                      type: CustomCollectionFieldType.custom,
                      onCustomRender: (field, value, onUpdate, item, itemId, onError) => {
                        return  React.createElement("div", {style: {minWidth: "150px"}}, React.createElement(
                          FieldPicker, 
                          {
                            context: this.context as any,
                            includeHidden: false,
                            includeReadOnly: false,
                            multiSelect: false,
                            orderBy: FieldsOrderBy.Title,
                            listId: this.properties.list,
                            selectedFields: value,
                            onSelectionChanged: (newvalue: ISPField ) => { onUpdate(field.id, newvalue.InternalName); },
                            showBlankOption: true,
                          }));
                      },
                      required: true
                    },
                    {
                      id: "width",
                      title: "Width",
                      type: CustomCollectionFieldType.number,
                      required: true
                    },
                    {
                      id: "groupByField",
                      title: "Group By Field (Only for first column)",
                      type: CustomCollectionFieldType.custom,
                      onCustomRender: (field, value, onUpdate, item, itemId, onError) => {
                        return React.createElement(
                          FieldPicker, 
                          {
                            context: this.context as any,
                            includeHidden: false,
                            includeReadOnly: false,
                            multiSelect: false,
                            orderBy: FieldsOrderBy.Title,
                            listId: this.properties.list,
                            selectedFields: value,
                            onSelectionChanged: (newvalue: ISPField ) => { onUpdate(field.id, newvalue.InternalName); },
                            showBlankOption: true
                          });
                      },
                      required: false
                    },
                  ],
                  disabled: false
                }),
                PropertyFieldColumnPicker('fullWidth', {
                  label: 'Full width column flag',
                  context: this.context as any,
                  selectedColumn: this.properties.fullWidth,
                  listId: this.properties.list,
                  disabled: false,
                  orderBy: PropertyFieldColumnPickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'fullWidth',
                  displayHiddenColumns: false,
                  columnReturnProperty: IColumnReturnProperty["Internal Name"]
              }),
              PropertyFieldColumnPicker('orderBy', {
                label: 'Order By',
                context: this.context as any,
                selectedColumn: this.properties.orderBy,
                listId: this.properties.list,
                disabled: false,
                orderBy: PropertyFieldColumnPickerOrderBy.Title,
                onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                properties: this.properties,
                onGetErrorMessage: null,
                deferredValidationTime: 0,
                key: 'orderBy',
                displayHiddenColumns: false,
                columnReturnProperty: IColumnReturnProperty["Internal Name"]
              }),
              PropertyFieldCodeEditor('footer', {
                label: 'Edit Footer HTML Code',
                panelTitle: 'Edit Footer HTML Code',
                initialValue: this.properties.footer,
                onPropertyChange: this.onPropertyPaneFieldChanged,
                properties: this.properties,
                disabled: false,
                key: 'footer',
                language: PropertyFieldCodeEditorLanguages.HTML,
                options: {
                  wrap: true,
                  fontSize: 20,
                  // more options
                }
              })
              ]
            }
          ]
        }
      ]
    };
  }
}
