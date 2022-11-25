import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './PnpcrudWebPart.module.scss';
import * as strings from 'PnpcrudWebPartStrings';
import * as pnp from 'sp-pnp-js';

export interface IPnpcrudWebPartProps {
  description: string;
}

export default class PnpcrudWebPart extends BaseClientSideWebPart<IPnpcrudWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

 
  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.pnpcrud} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
      </div>
      <div>
       ID: <input type='text' id='txtId'/><br/><br/>
       <input type='button' id='btnRead' value='Read details'/><br/><br/>
       Title: <input type='text' id='txtSoftwareTitle'/><br/><br/>
      
      <input type='button' id='btnAddListItem' value='Add new item'/><br/><br/>
      <input type='button' id='btnDelete' value='Delete'/><br/><br/>
      <input type='button' id='btnUpdate' value='Update'/><br/><br/>
      <input type='button' id='btnShowAll' value='Show all'/><br/><br/>
  </div>
  <div id="divStatus">
  </div
    </section>`;
    this.bindElements();
  }


  private bindElements(): void {
    this.domElement.querySelector("#btnAddListItem").addEventListener('click', ()=>{this.AddListItem()});
    this.domElement.querySelector("#btnUpdate").addEventListener('click', ()=>{this.UpdateListItem()});
    this.domElement.querySelector("#btnRead").addEventListener('click', ()=>{this.ReadListItem()});
    this.domElement.querySelector("#btnShowAll").addEventListener('click', ()=>{this.ReadAllListItems()});
    this.domElement.querySelector("#btnDelete").addEventListener('click', ()=>{this.DeleteItem()});
  }

   private ReadListItem(): void {
    let id : number;
    id = + (document.getElementById("txtId") as HTMLInputElement).value;
    pnp.sp.web.lists.getByTitle("SoftwareCatalog").items.getById(id).get().then((item:any) => {
      const input = (document.getElementById("txtSoftwareTitle")  as HTMLInputElement).value = item.Title;
    });
  }



  private ReadAllListItems(): void {
    let html : string = "<table><th>ID</th><th>Title</th>";

    pnp.sp.web.lists.getByTitle("SoftwareCatalog").items.get().then((items:any[])=>{
      items.forEach(item => {
        html+=`<tr><td>${item.ID}</td><td>${item.Title}</td></tr>`
        
      });
      html+=`</table>`;
      const listContainer : Element = this.domElement.querySelector("#divStatus")
      listContainer.innerHTML = html;

    });
  }

  
  private AddListItem(): void {
    let txtSoftwareTitle = (document.getElementById("txtSoftwareTitle") as HTMLInputElement).value;

    pnp.sp.web.lists.getByTitle("SoftwareCatalog").items.add({
        Title: txtSoftwareTitle
    }).then(response=> {
      alert("success");
    });
      
      this.clear();
    
  }

  private UpdateListItem(): void {
    let txtSoftwareTitle = (document.getElementById("txtSoftwareTitle") as HTMLInputElement).value;
    let id: number = + (document.getElementById("txtId") as HTMLInputElement).value;

    pnp.sp.web.lists.getByTitle("SoftwareCatalog").items.getById(id).update({
      Title: txtSoftwareTitle
    }).then(Response=>{
      alert("updated");
    });
  }

  private DeleteItem(): void {
    let id : number;
    id = + (document.getElementById("txtId") as HTMLInputElement).value;
    pnp.sp.web.lists.getByTitle("SoftwareCatalog").items.getById(id).delete().then((item:any) => {
      alert("deleted")
    });
  }

  private clear(): void {
    (document.getElementById("txtSoftwareTitle") as HTMLInputElement).value = "";
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
