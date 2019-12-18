import { Component } from '@angular/core';
import * as XLSX from 'xlsx';
import * as FileSaver from 'file-saver';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'import-xls';

  storeData: any;
  csvData: any;
  jsonData: any = [];
  textData: any;
  htmlData: any;
  fileUploaded: File;
  jsonWorkSheet: any = [];
  worksheet: any;

  uploadedFile(event) {
    this.fileUploaded = event.target.files[0];
    this.readExcel();
  }

  readExcel() {
    const readFile = new FileReader();
    readFile.onload = (e) => {
      this.storeData = readFile.result;
      const data = new Uint8Array(this.storeData);
      const arr = new Array();
      for (let i = 0; i !== data.length; ++i) { arr[i] = String.fromCharCode(data[i]); }
      const bstr = arr.join('');
      const workbook = XLSX.read(bstr, { type: 'binary' });
      const SHEET_NAME: string[] =  workbook.SheetNames;
      this.worksheet = workbook.Sheets[SHEET_NAME[0]];
      for (let i = 0; i !== SHEET_NAME.length; i++ ) {
        this.jsonWorkSheet.push( {sheetNames:SHEET_NAME[i] , sheets: workbook.Sheets[SHEET_NAME[i]]});
      }
    };
    readFile.readAsArrayBuffer(this.fileUploaded);
  }

  readAsCSV() {
    this.csvData = XLSX.utils.sheet_to_csv(this.worksheet);
    const data: Blob = new Blob([this.csvData], { type: 'text/csv;charset=utf-8;' });
    FileSaver.saveAs(data, 'CSVFile' + new Date().getTime() + '.csv');
  }

  readAsJson() {

    for(let i = 0; i !== this.jsonWorkSheet.length; i++ ) {
      this.jsonData.push( {
        sheetNames: this.jsonWorkSheet[i].sheetNames,
        sheets:  XLSX.utils.sheet_to_json(this.jsonWorkSheet[i].sheets, { raw: false })
      });
    }
    this.jsonData = JSON.stringify(this.jsonData);
    const data: Blob = new Blob([this.jsonData], { type: 'application/json' });
    FileSaver.saveAs(data, 'JsonFile' + new Date().getTime() + '.json');
  }

  readAsHTML() {
    this.htmlData = XLSX.utils.sheet_to_html(this.worksheet);
    const data: Blob = new Blob([this.htmlData], { type: 'text/html;charset=utf-8;' });
    FileSaver.saveAs(data, 'HtmlFile' + new Date().getTime() + '.html');
  }

  readAsText() {
    this.textData = XLSX.utils.sheet_to_txt(this.worksheet);
    const data: Blob = new Blob([this.textData], { type: 'text/plain;charset=utf-8;' });
    FileSaver.saveAs(data, 'TextFile' + new Date().getTime() + '.txt');
  }

}
