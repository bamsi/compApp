import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import * as XLSX from 'xlsx';
import { findBestMatch } from 'string-similarity';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'compApp';

  // File 1 data
  data1: any[] = [];
  headers1: string[] = [];
  selectedColumn1: string = '';

  // File 2 data
  data2: any[] = [];
  headers2: string[] = [];
  selectedColumn2: string = '';

  // Results
  results: any[] = [];

  onFileChange(evt: any, fileNumber: number) {
    const target: DataTransfer = <DataTransfer>(evt.target);
    if (target.files.length !== 1) {
      throw new Error('Cannot use multiple files');
    }
    const reader: FileReader = new FileReader();
    reader.onload = (e: any) => {
      const bstr: string = e.target.result;
      const wb: XLSX.WorkBook = XLSX.read(bstr, { type: 'binary' });
      const wsname: string = wb.SheetNames[0];
      const ws: XLSX.WorkSheet = wb.Sheets[wsname];
      const data = <any[]>XLSX.utils.sheet_to_json(ws, { header: 1 });

      if (fileNumber === 1) {
        this.headers1 = data[0];
        this.data1 = data.slice(1).map((arr: any[]) => {
          const row: any = {};
          this.headers1.forEach((header, index) => {
            row[header] = arr[index];
          });
          return row;
        });
      } else {
        this.headers2 = data[0];
        this.data2 = data.slice(1).map((arr: any[]) => {
          const row: any = {};
          this.headers2.forEach((header, index) => {
            row[header] = arr[index];
          });
          return row;
        });
      }
    };
    reader.readAsBinaryString(target.files[0]);
  }

  compareFiles() {
    this.results = [];
    const targetStrings = this.data2.map(row => String(row[this.selectedColumn2] || ''));

    this.data1.forEach(row1 => {
      const sourceString = String(row1[this.selectedColumn1] || '');
      if (sourceString && targetStrings.length > 0) {
        const bestMatch = findBestMatch(sourceString, targetStrings);
        if (bestMatch.bestMatch.rating > 0) { // Consider adding a confidence threshold
          const file2Row = this.data2[bestMatch.bestMatchIndex];
          this.results.push({
            file1Row: row1,
            file2Row: file2Row,
            confidence: bestMatch.bestMatch.rating
          });
        }
      }
    });
  }

  downloadResults() {
    const dataToExport = this.results.map(result => {
      const row: any = {};
      // Add data from file 1
      for (const header of this.headers1) {
        row[header] = result.file1Row[header];
      }
      // Add data from file 2 with a suffix
      for (const header of this.headers2) {
        row[`${header} (Matched)`] = result.file2Row ? result.file2Row[header] : '';
      }
      // Add confidence score
      row['Confidence'] = result.confidence;
      return row;
    });

    const ws: XLSX.WorkSheet = XLSX.utils.json_to_sheet(dataToExport);
    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Results');
    XLSX.writeFile(wb, 'comparison_results.csv');
  }
}
