import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'LeaverequestWebPartStrings';
import Leaverequest from './components/Leaverequest';
import { ILeaverequestProps } from './components/ILeaverequestProps';

import { sp } from "@pnp/sp/preset/all";

import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls';

import {
  ThemeProvider,
  ThemeChangedEventArgs,
  IReadonlyTheme,
} from 'office-ui-fabric-react/lib/Foundation';

export interface ILeaverequestWebPartProps {
  description: string;
  storageList: string;
  acknowledgementMessage: string;
  readMessage: string;
}

export default class LeaverequestWebPart extends BaseClientSideWebPart<ILeaverequestWebPartProps> {

  private _themeProvider: ThemeProvider;
  private _themeVariant: IReadonlyTheme | undefined;

  protected async onInit(): Promise<void> {
    await super.onInit();

    sp.setup(this.context);

    this._themeProvider = this.context.serviceScope.consume(
      ThemeProvider.serviceKey
    );

    this.__themeVariant = this._themeProvider.tryGetTheme();
    this.onThemeChangedEvent.add(
      this,
      this._handleThemeChangedEvent
    );
  }
  private _handleThemeChangedEvents(args: ThemeChangedEventArgs): void {


    this._themeVariant = args.theme;
    this.render();
  }
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<ILeaverequestProps> = React.createElement(
      Leaverequest,
      {
        documentTitle: this.properties.documentTitle,
        currentUserDisplayName: this.context.pageContext.user.displayName,
        storagelist: this.properties.storageList,
        acknowledgementLabel: this.properties.acknowledgementlabel,
        acknowledgmentMessage: this.properties.acknowledgementMessage,
        readMessage: this.properties.readMessage,
        themeVariant: this, themeVariant,
        configured: this.properties.storageList ? this.properties.storageList != '' : false,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

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
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
