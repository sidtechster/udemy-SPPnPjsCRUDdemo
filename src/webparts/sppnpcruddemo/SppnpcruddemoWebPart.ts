import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SppnpcruddemoWebPart.module.scss';
import * as strings from 'SppnpcruddemoWebPartStrings';

import * as pnp from 'sp-pnp-js';

export interface ISppnpcruddemoWebPartProps {
  description: string;
}

export default class SppnpcruddemoWebPart extends BaseClientSideWebPart<ISppnpcruddemoWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.sppnpcruddemo }">
      <p>Enter ID</p><br/>
      <input type="text" id="txtID"/>
      <input type="submit" value="Read details" id="btnRead" />
      <br/><br/>

      <p>Title</p><br/>
      <input type="text" id="txtTitle"/><br/><br/>
      
      <input type="submit" value="Insert item" id="btnSubmit" />
      <input type="submit" value="Update item" id="btnUpdate" />
      <input type="submit" value="Delete item" id="btnDelete" />
      <input type="submit" value="Show all item" id="btnShowAll" />

      <br/>

      <div id="divStatus"></div>

      <div id="spListData"></div>
      </div>`;

      this.bindEvents();
      this.readAllItems();
  }

  private bindEvents(): void {
    this.domElement.querySelector('#btnSubmit').addEventListener('click', () => { this.addListItem(); });
    this.domElement.querySelector('#btnRead').addEventListener('click', () => { this.readListItem(); });
    this.domElement.querySelector('#btnUpdate').addEventListener('click', () => { this.updateListItem(); });
    this.domElement.querySelector('#btnDelete').addEventListener('click', () => { this.deleteListItem(); });
    this.domElement.querySelector('#btnShowAll').addEventListener('click', () => { this.readAllItems(); });
  }

  private readAllItems(): void {
    let html: string = '<h2>Title</h2>';
    html += '<ul>';

    pnp.sp.web.lists.getByTitle("MySampleList").items.get().then((items: any[]) => {
      items.forEach(function(item) {
        html += `<li>${item["Title"]}</li>`;
      });

      html += '</ul>';
      const listContainer: Element = this.domElement.querySelector('#spListData');
      listContainer.innerHTML = html;
    });
  }

  private deleteListItem(): void {
    let id = document.getElementById('txtID')["value"];

    pnp.sp.web.lists.getByTitle("MySampleList").items.getById(id).delete()
      .then(r => {
        alert("Deleted");
      })
  }

  private updateListItem(): void {
    let id = document.getElementById('txtID')["value"];
    let title = document.getElementById('txtTitle')["value"];

    pnp.sp.web.lists.getByTitle("MySampleList").items.getById(id).update({
      Title: title
    }).then(r => {
      alert("Update success");
    });
  }

  private readListItem(): void {
    let id = document.getElementById('txtID')["value"];

    pnp.sp.web.lists.getByTitle("MySampleList").items.getById(id).get().then((item: any) => {
      document.getElementById("txtTitle")["value"] = item["Title"];
    });
  }

  private addListItem(): void {
    let title = document.getElementById('txtTitle')["value"];

    pnp.sp.web.lists.getByTitle("MySampleList").items.add({
      Title: title
    }).then(r => {
      alert("Success");
    });
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
