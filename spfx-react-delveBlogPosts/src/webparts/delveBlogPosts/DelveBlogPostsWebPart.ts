import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneToggle
} from '@microsoft/sp-webpart-base';

import * as strings from 'DelveBlogPostsWebPartStrings';
import DelveBlogPosts from './components/DelveBlogPosts';
import { IDelveBlogPostsProps } from './components/IDelveBlogPostsProps';

import { sp } from "@pnp/sp";
import { PropertyFieldPeoplePicker, PrincipalType, IPropertyFieldGroupOrPerson  } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';

export interface IDelveBlogPostsWebPartProps {
  rowLimit: number;
  commandBar: boolean;
  people: IPropertyFieldGroupOrPerson[];
}

export default class DelveBlogPostsWebPart extends BaseClientSideWebPart<IDelveBlogPostsWebPartProps> {

  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      // init @pnp/sp to access it on render
      sp.setup({
        spfxContext: this.context
      });
    });
  }

  public render(): void {
    const element: React.ReactElement<IDelveBlogPostsProps > = React.createElement(
      DelveBlogPosts,
      {
        rowLimit: this.properties.rowLimit,
        commandBar: this.properties.commandBar,
        context: this.context,
        people: this.properties.people
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
          displayGroupsAsAccordion: true,
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.LayoutGroupName,
              groupFields: [
                PropertyPaneSlider('rowLimit', {
                  label:"Max Items",
                  min:1,
                  max:10,
                  value:3,
                  showValue:true,
                  step:1
                }),
                PropertyPaneToggle('commandBar', {
                  label:"Show commands",
                  checked:true
                })
              ]
            },
            {
              groupName: strings.FilterGroupName, 
              groupFields: [
                PropertyFieldPeoplePicker('people', {
                  label:"Select specific user(s)",
                  allowDuplicate:false,
                  context:this.context,
                  properties:this.properties,
                  key:'peopleId',
                  multiSelect:true,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
