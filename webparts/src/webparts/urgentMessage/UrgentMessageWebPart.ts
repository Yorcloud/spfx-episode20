import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'UrgentMessageWebPartStrings';
import UrgentMessage from './components/UrgentMessage';
import { IUrgentMessageProps } from './components/IUrgentMessageProps';

import { sp } from "@pnp/sp/presets/all";

import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';

import {
  ThemeProvider,
  ThemeChangedEventArgs,
  IReadonlyTheme,
} from "@microsoft/sp-component-base";

// List Guid = ad499396-0284-473b-af73-5144967eef93

export interface IUrgentMessageWebPartProps {
  description: string;
  list: string;
  label: string;
  message: string;
}

export default class UrgentMessageWebPart extends BaseClientSideWebPart<IUrgentMessageWebPartProps> {

  private _themeProvider: ThemeProvider;
  private _themeVariant: IReadonlyTheme | undefined;

  protected async onInit(): Promise<void> {

    await super.onInit();

    sp.setup(this.context);

    this._themeProvider = this.context.serviceScope.consume(
      ThemeProvider.serviceKey
    );

    // If it exists, get the theme variant
    this._themeVariant = this._themeProvider.tryGetTheme();

    // Register a handler to be notified if the theme variant changes
    this._themeProvider.themeChangedEvent.add(
      this,
      this._handleThemeChangedEvent
    );



  }

  private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme;
    this.render();
  }

  public render(): void {
    const element: React.ReactElement<IUrgentMessageProps> = React.createElement(
      UrgentMessage,
      {
        description: this.properties.description,
        currentUser: this.context.pageContext.user.loginName,
        list: this.properties.list,
        label: this.properties.label,
        message: this.properties.message,
        themeVariant: this._themeVariant
      }
    );

    ReactDom.render(element, this.domElement);
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
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyFieldListPicker('list', {
                  label: 'Select a list',
                  selectedList: this.properties.list,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId',
                  multiSelect: false
                }),
                PropertyPaneTextField('label', {
                  label: "Label for Checkbox"
                }),
                PropertyPaneTextField('message', {
                  label: "Message to display"
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
