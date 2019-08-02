import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './PpControlsTestWebPart.module.scss';
import * as strings from 'PpControlsTestWebPartStrings';

import {
  PropertyFieldCollectionData,
  CustomCollectionFieldType
} from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import {
IPropertyFieldGroupOrPerson,
PropertyFieldPeoplePicker,
PrincipalType
} from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';

export interface IPpControlsTestWebPartProps {
  description: string;
  people: IPropertyFieldGroupOrPerson[];
  expansionOpions: any[];
}

export default class PpControlsTestWebPart extends BaseClientSideWebPart<IPpControlsTestWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.ppControlsTest }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <p class="${ styles.description }">${this.properties.people}</p>
              <p class="${ styles.description }">${this.properties.expansionOpions}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;
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
                PropertyFieldPeoplePicker('people', {
                label: 'PP Field People Picker Control',
                initialData: this.properties.people,
                allowDuplicate: false,
                principalType: [PrincipalType.Users, PrincipalType.SharePoint, PrincipalType.Security],
                onPropertyChange: this.onPropertyPaneFieldChanged,
                context: this.context,
                properties: this.properties,
                onGetErrorMessage: null,
                deferredValidationTime: 0,
                key: 'peopleFieldId'
                }),
                PropertyFieldCollectionData('expansionOptions', {
                  key: 'collectionData',
                  label: 'Possible expansion options',
                  panelHeader: 'Possible expansion options',
                  manageBtnLabel: 'Manage expansion options',
                  value: this.properties.expansionOpions,
                  fields: [
                    {
                      id: 'Region',
                      title: 'Region',
                      required: true,
                      type: CustomCollectionFieldType.dropdown,
                      options: [
                        
                          { key: 'WA', text: 'Western Australia' },
                          { key: 'VIC', text: 'Victoria' },
                          { key: 'NSW', text: 'New South Wales' },
                          { key: 'QLD', text:'Queensland' },
                          { key: 'SA', text: 'South Australia' },
                          { key: 'TAS', text: 'Tasmania' },
                          { key: 'NT', text: 'Northern Territory' },
                          { key: 'ACT', text: 'Australian Capital Territory' }
                      ]
                    },
                    {
                      id: 'Comment',
                      title: 'Comment',
                      type: CustomCollectionFieldType.string
                    }
                  ]
                })                        
              ]
            }
          ]
        }
      ]
    };
  }
}