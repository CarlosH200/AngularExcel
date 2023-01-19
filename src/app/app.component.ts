import { Component } from '@angular/core';
import * as GC from '@grapecity/spread-sheets';
import * as Excel from '@grapecity/spread-excelio';
import { SpreadSheetsModule } from '@grapecity/spread-sheets-angular';
import { saveAs } from 'file-saver';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'AngularEXCEL';

  private spread;
  private excelIO;

  constructor() {
    this.spread = new GC.Spread.Sheets.Workbook();
    this.excelIO = new Excel.IO();
  }

  arrayDataJson: any[] = []

  // Funcion que le el excel  lo guarda en una variable Vaiable: json
  onFileChange(args: any) {
    const self = this, file = args.srcElement && args.srcElement.files && args.srcElement.files[0];
    if (self.spread && file) {
      self.excelIO.open(file, (json: any) => {
        self.spread.fromJSON(json, {});
        setTimeout(() => {
          alert('ARCHIVO MIGRADO');
          // console.log(json);
          // console.log(json.sheets.Hoja1.data.dataTable[13][0]);


          // console.log(Object.values(json.sheets.Hoja1.data.dataTable));

          // Saca los valores de la propiedad y lo asigna a un arreglo (arrayRows)
          let arrayRows = Object.values(json.sheets.Hoja1.data.dataTable)

          arrayRows.forEach(element => {
            let columnArray = Object.values(<Object>element)

            columnArray.forEach(columnElement => {
              this.arrayDataJson.push(columnElement);

            });

          });



        }, 0);
      }, (error: any) => {
        // alert('ERROR AL MIGRAR ARCHIVO');
      });
    }
  }
  // onClickMe(args: any) {
  //   const self = this;
  //   const filename = 'exportExcel.xlsx';
  //   const json = JSON.stringify(self.spread.toJSON());
  //   self.excelIO.save(json, function (blob: any) {
  //     saveAs(blob, filename);
  //   }, function (error: any) {
  //       console.log(error);
  //   });
  // }


}
