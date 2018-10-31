import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './AnnouncementsWebPart.module.scss';
import * as strings from 'AnnouncementsWebPartStrings';
import { SPComponentLoader } from '@microsoft/sp-loader';
import 'jquery';
require('bootstrap');

export interface IAnnouncementsWebPartProps {
  description: string;
}

export default class AnnouncementsWebPart extends BaseClientSideWebPart<IAnnouncementsWebPartProps> {

  public render(): void {
    let url = "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";
    SPComponentLoader.loadCss(url);
    this.domElement.innerHTML = `
            
    <div id="myCarousel" class="carousel slide" data-ride="carousel">
    
      <!-- Wrapper for slides -->
      <div class="carousel-inner">

      <div class='item active'><img src="data:image/jpeg;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/2wCEAAkGBxISEhEQEBIWFhUXFxcaFxcVGBcXFRgYGBcWGxgZFRUaHikgGhwlHRcVITEhJSkrLi4vGR8zODMsNygtLjcBCgoKDQ0NDg0NDisZFRkrNys3KystKysrKysrKysrKysrKysrKysrKysrKysrKysrKysrKysrKysrKysrKysrK//AABEIALcBEwMBIgACEQEDEQH/xAAcAAEAAgIDAQAAAAAAAAAAAAAABgcFCAECBAP/xABAEAABAwIEAwYDAwoFBQAAAAABAAIRAyEEBRIxBkFRBxMiYXGBMpGhFLHwIzNCQ1JicoLB8QgVkqLhJFNjs9H/xAAUAQEAAAAAAAAAAAAAAAAAAAAA/8QAFBEBAAAAAAAAAAAAAAAAAAAAAP/aAAwDAQACEQMRAD8AvFERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBEXBKDlFGM64+y7C2q4qmXfssIe7obN9/koZW7bqAc4so6mXiX6XkA2JBEAm9vryQW0irDL+2/L3kNr061Hq4tFRnzYSfop9k2e4XFt14WvTqjnocCR/E3dvuEGRREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBEWC414kp5fhKmJqESARTaTd9Qg6Wj5SegBKD48acZ4XLKQqYhxLnT3dNt3vI6cg0SJcbCfZa+cYdoWOzCW1Hd3SvFGkSAQZ/Ou3fYxe3ko/jcbXx2JNaq81KryLmT6ADkN7CAFLsHwhoZrI1OEzPXnby/oghVLDON9P4/ELscveb+alv+WuJaGjcGBtsF925boA1sOl3Pl/xdBBauCIXNA1KLm1aT303tuH0yWuB6hwuFM6mUartNufleL/VfLEZPIIIugnXZt2wa3MwmaEBxhrMQBDXE2ArDZp/eFusbq5wVppmeXmmfL8bq2OxDtAqd43K8W8ua62He65aRJNNzuYI+GdtuYgL0REQEREBERAREQEREBERAREQEREBERAREQEREBERAREQFq92t8Xf5hjSym6cPQJZTgyHEHx1ByubA9AFsBx/m4wmAxVctLvAWgA6bv8ACPELgCZkXWpjWbRytfyhBZ3ZfkgDXYhzBJMMJi0bkeSntfCAmY9fNfLh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" alt="eonMusk" style='width:100%;height:500px;opacity: 0.8;'>
      <div style=" position: absolute;top: 10%;left: 20%;transform: translate(-50%, -50%);
      font-size: 24px;color:black;">Steve jobs</div>
      <div style=" position: absolute;top: 20%;left:40%;transform: translate(-50%, -50%);
      font-size:12px;color:black;display:block;">He was the chairman,chief executive officer(CEO) and co-founder of Apple Inc</div>
      </div>
      
      <div class="item"><img src="https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQOlfeU7AcmH9xevmakiz8uGQaK_e31zNRnEowAjKuAKaMb9yDH" alt="eonMusk" style='width:100%;height:500px;opacity: 0.8;'>
      <div style=" position: absolute;top: 10%;left: 20%;transform: translate(-50%, -50%);
      font-size: 24px;color:black;">Elon Musk</div>
      <div style=" position: absolute;top: 20%;left: 40%;transform: translate(-50%, -50%);
      font-size:12px;color:black;display:block;">Founder, CEO, and lead designer of SpaceX, Co-founder, CEO, and product architect of Tesla Inc</div>
      </div>
	  
	    <div class="item"><img src="https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcSLzu_o4lHY2oEaW_5Lyni-F3OoHrVAs3uXT3P60RtElGlOrkHLEg" alt="eonMusk" style='width:100%;height:500px;opacity: 0.8;'>
      <div style=" position: absolute;top: 10%;left: 20%;transform: translate(-50%, -50%);
      font-size: 24px;color:black;">Ratan tata</div>
      <div style=" position: absolute;top: 20%;left: 40%;transform: translate(-50%, -50%);
      font-size:12px;color:black;display:block;">Indian industrialist, investor, philanthropist, and former chairman of Tata Sons and Tata Group</div>
      </div>
      
      <div class='item'><img src="data:image/jpeg;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/2wCEAAkGBxISEhEQEBIWFhUXFxcaFxcVGBcXFRgYGBcWGxgZFRUaHikgGhwlHRcVITEhJSkrLi4vGR8zODMsNygtLjcBCgoKDQ0NDg0NDisZFRkrNys3KystKysrKysrKysrKysrKysrKysrKysrKysrKysrKysrKysrKysrKysrKysrK//AABEIALcBEwMBIgACEQEDEQH/xAAcAAEAAgIDAQAAAAAAAAAAAAAABgcFCAECBAP/xABAEAABAwIEAwYDAwoFBQAAAAABAAIRAyEEBRIxBkFRBxMiYXGBMpGhFLHwIzNCQ1JicoLB8QgVkqLhJFNjs9H/xAAUAQEAAAAAAAAAAAAAAAAAAAAA/8QAFBEBAAAAAAAAAAAAAAAAAAAAAP/aAAwDAQACEQMRAD8AvFERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBEXBKDlFGM64+y7C2q4qmXfssIe7obN9/koZW7bqAc4so6mXiX6XkA2JBEAm9vryQW0irDL+2/L3kNr061Hq4tFRnzYSfop9k2e4XFt14WvTqjnocCR/E3dvuEGRREQEREBERAREQEREBERAREQEREBERAREQEREBERAREQEREBEWC414kp5fhKmJqESARTaTd9Qg6Wj5SegBKD48acZ4XLKQqYhxLnT3dNt3vI6cg0SJcbCfZa+cYdoWOzCW1Hd3SvFGkSAQZ/Ou3fYxe3ko/jcbXx2JNaq81KryLmT6ADkN7CAFLsHwhoZrI1OEzPXnby/oghVLDON9P4/ELscveb+alv+WuJaGjcGBtsF925boA1sOl3Pl/xdBBauCIXNA1KLm1aT303tuH0yWuB6hwuFM6mUartNufleL/VfLEZPIIIugnXZt2wa3MwmaEBxhrMQBDXE2ArDZp/eFusbq5wVppmeXmmfL8bq2OxDtAqd43K8W8ua62He65aRJNNzuYI+GdtuYgL0REQEREBERAREQEREBERAREQEREBERAREQEREBERAREQFq92t8Xf5hjSym6cPQJZTgyHEHx1ByubA9AFsBx/m4wmAxVctLvAWgA6bv8ACPELgCZkXWpjWbRytfyhBZ3ZfkgDXYhzBJMMJi0bkeSntfCAmY9fNfLh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" alt="eonMusk" style='width:100%;height:500px;opacity: 0.8;'>
      <div style=" position: absolute;top: 10%;left: 20%;transform: translate(-50%, -50%);
      font-size: 24px;color:black;">Steve jobs</div>
      <div style=" position: absolute;top: 20%;left:40%;transform: translate(-50%, -50%);
      font-size:12px;color:black;display:block;">He was the chairman,chief executive officer(CEO) and co-founder of Apple Inc</div>
      </div>

      <div class="item"><img src="https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQOlfeU7AcmH9xevmakiz8uGQaK_e31zNRnEowAjKuAKaMb9yDH" alt="eonMusk" style='width:100%;height:500px;opacity: 0.8;'>
      <div style=" position: absolute;top: 10%;left: 20%;transform: translate(-50%, -50%);
      font-size: 24px;color:black;">Elon Musk</div>
      <div style=" position: absolute;top: 20%;left: 40%;transform: translate(-50%, -50%);
      font-size:12px;color:black;display:block;">Founder, CEO, and lead designer of SpaceX, Co-founder, CEO, and product architect of Tesla Inc</div>
      </div>

      </div>
</div>

  <!-- Left and right controls -->
        <button style="background-color:yellow" type="button" class="btn btn-default btn-sm">
        <span class="glyphicon glyphicon-chevron-left" href="#myCarousel" data-slide="prev"></span>
        </button>
        <button style="background-color:yellow" type="button" class="btn btn-default btn-sm">
        <span class="glyphicon glyphicon-chevron-right" href="#myCarousel" data-slide="next"></span>
        </button>
      `;
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
