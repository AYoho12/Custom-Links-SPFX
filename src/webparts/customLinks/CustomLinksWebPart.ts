import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneButton,
  PropertyPaneButtonType,
  IPropertyPaneField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './CustomLinksWebPart.module.scss';
import * as strings from 'CustomLinksWebPartStrings';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";

export interface ICustomLinksWebPartProps {
  description: string;
  linkTitle: string;
  linkUrl: string;
}

export default class CustomLinksWebPart extends BaseClientSideWebPart<ICustomLinksWebPartProps> {
  private links: { title: string, url: string }[] = [];
  private editingLinkIndex: number | null = null;
  private userId: number;

  public async onInit(): Promise<void> {
    await super.onInit();
    sp.setup({
      spfxContext: this.context as any
    });

    // Get the current user's ID
    const currentUser = await sp.web.currentUser.get();
    this.userId = currentUser.Id;

    // Load the user's links
    await this.loadUserLinks();
  }

  private async loadUserLinks(): Promise<void> {
    const items = await sp.web.lists.getByTitle('UserLinks').items
      .filter(`UserId eq ${this.userId}`)
      .select('Id', 'Title', 'Url')
      .get();

    this.links = items.map(item => ({ title: item.Title, url: item.Url }));
    this.render();
  }

  private async saveUserLinks(): Promise<void> {
    // Clear existing links for the user
    const items = await sp.web.lists.getByTitle('UserLinks').items
      .filter(`UserId eq ${this.userId}`)
      .get();
      
    for (const item of items) {
      await sp.web.lists.getByTitle('UserLinks').items.getById(item.Id).delete();
    }

    // Add new links
    for (const link of this.links) {
      await sp.web.lists.getByTitle('UserLinks').items.add({
        Title: link.title,
        Url: link.url,
        UserId: this.userId
      });
    }
  }

  private async deleteLink(): Promise<void> {
    if (this.editingLinkIndex !== null) {
      const linkToDelete = this.links[this.editingLinkIndex];
      const items = await sp.web.lists.getByTitle('UserLinks').items
        .filter(`Title eq '${linkToDelete.title}' and Url eq '${linkToDelete.url}' and UserId eq ${this.userId}`)
        .get();
      
      for (const item of items) {
        await sp.web.lists.getByTitle('UserLinks').items.getById(item.Id).delete();
      }

      this.links.splice(this.editingLinkIndex, 1);
      this.editingLinkIndex = null;
      await this.saveUserLinks();
      this.render();
    }
  }

  public render(): void {
    let linksHtml: string = '';
    if (this.links.length > 0) {
      linksHtml = this.links.map((link, index) => `
        <tr>
          <td class="${styles.linkFormat}"><a href="${escape(link.url)}" target="_blank">${escape(link.title)}</a></td>
          <td><button class="${styles.editButton}" data-index="${index}">Edit</button></td>
        </tr>
      `).join('');
    }

    this.domElement.innerHTML = `
    <section class="${styles.customLinks} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <button class="${styles.addButton}" id="addLinkButton">Add Link</button>
      <table>
        <thead>
          <tr>
            <th>Link</th>
          </tr>
        </thead>
        <tbody>
          ${linksHtml || '<tr><td colspan="2">No links added yet</td></tr>'}
        </tbody>
      </table>
    </section>`;

    // Add click event listeners to the "Add Link" and "Edit" buttons
    this.bindEvents();
  }

  private bindEvents(): void {
    const addButton: HTMLElement = this.domElement.querySelector('#addLinkButton')!;
    if (addButton) {
      addButton.addEventListener('click', () => {
        this.editingLinkIndex = null; // Reset the editing index for a new link
        this.context.propertyPane.open();
      });
    }

    const editButtons: NodeListOf<HTMLButtonElement> = this.domElement.querySelectorAll(`.${styles.editButton}`);
    editButtons.forEach(button => {
      button.addEventListener('click', (event) => this.onEditButtonClick(event));
    });
  }

  private onEditButtonClick(event: Event): void {
    const target = event.target as HTMLButtonElement;
    const index = parseInt(target.getAttribute('data-index')!, 10);
    this.editingLinkIndex = index;
    this.properties.linkTitle = this.links[index].title;
    this.properties.linkUrl = this.links[index].url;
    this.context.propertyPane.open();
  }

  private async addLink(): Promise<void> {
    if (this.properties.linkTitle && this.properties.linkUrl) {
      if (this.editingLinkIndex !== null) {
        this.links[this.editingLinkIndex] = {
          title: this.properties.linkTitle,
          url: this.properties.linkUrl
        };
        this.editingLinkIndex = null;
      } else {
        this.links.push({
          title: this.properties.linkTitle,
          url: this.properties.linkUrl
        });
      }
      await this.saveUserLinks();
      this.render();
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const buttonText = this.editingLinkIndex !== null ? 'Update' : 'Add';
  
    const fields: IPropertyPaneField<any>[] = [
      PropertyPaneTextField('linkTitle', {
        label: 'Link Title'
      }),
      PropertyPaneTextField('linkUrl', {
        label: 'Link URL'
      }),
      PropertyPaneButton('addLink', {
        text: buttonText,
        buttonType: PropertyPaneButtonType.Primary,
        onClick: this.addLink.bind(this)
      })
    ];
  
    if (this.editingLinkIndex !== null) {
      fields.push(PropertyPaneButton('deleteLink', {
        text: 'Delete',
        buttonType: PropertyPaneButtonType.Normal, // Use Primary and style it as Danger if needed
        onClick: this.deleteLink.bind(this)
      }));
    }
  
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: fields
            }
          ]
        }
      ]
    };
  }
  
}
