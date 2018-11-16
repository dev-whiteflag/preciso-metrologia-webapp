import { Component, OnInit } from '@angular/core';
import { Observable } from 'rxjs';
import 'rxjs/add/operator/map';

import { AngularFirestore, AngularFirestoreCollection, AngularFirestoreDocument } from 'angularfire2/firestore';
import { MatTableDataSource } from '@angular/material';
import { SelectionModel } from '@angular/cdk/collections';
import * as XLSX from 'xlsx-populate';

interface Certificado {
  // Dados Gerais
  Empresa: string;
  'Endereço': string;
  'Cidade c/ Estado': string;
  Equipamento: string;
  'Tipo de Instrumento': string;
  Instrumento: string;

  // Informações Especificas
  Grandeza: string;
  'Inicio da Faixa de Uso': string;
  'Final da Faixa de Uso': string;
  'Valor de uma Divisão': string;

  // Informações do Instrumento
  'Marca do Instrumento': string;
  'Modelo do Instrumento': string;
  'Classe do Instrumento': string;
  'Número de Série do Instrumento': string;
  'Identificação do Instrumento': string;

  //Dados Brutos
  'V.C.C 1 - Primeira Leitura': string;
  'V.C.C 1 - Incerteza - Primeira Leitura': string;
  'V.C.C 2 - Primeira Leitura': string;
  'V.C.C 2 - Incerteza - Primeira Leitura': string;
  'V.C.C 3 - Primeira Leitura': string;
  'V.C.C 3 - Incerteza - Primeira Leitura': string;
  'V.C.C 4 - Primeira Leitura': string;
  'V.C.C 4 - Incerteza - Primeira Leitura': string;
  'V.I.I 1.1 - Primeira Leitura': string;
  'V.I.I 1.2 - Primeira Leitura': string;
  'V.I.I 1.3 - Primeira Leitura': string;
  'V.I.I 2.1 - Primeira Leitura': string;
  'V.I.I 2.2 - Primeira Leitura': string;
  'V.I.I 2.3 - Primeira Leitura': string;
  'V.I.I 3.1 - Primeira Leitura': string;
  'V.I.I 3.2 - Primeira Leitura': string;
  'V.I.I 3.3 - Primeira Leitura': string;
  'V.I.I 4.1 - Primeira Leitura': string;
  'V.I.I 4.2 - Primeira Leitura': string;
  'V.I.I 4.3 - Primeira Leitura': string;

  'V.C.C 1 - Segunda Leitura': string;
  'V.C.C 1 - Incerteza - Segunda Leitura': string;
  'V.C.C 2 - Segunda Leitura': string;
  'V.C.C 2 - Incerteza - Segunda Leitura': string;
  'V.C.C 3 - Segunda Leitura': string;
  'V.C.C 3 - Incerteza - Segunda Leitura': string;
  'V.C.C 4 - Segunda Leitura': string;
  'V.C.C 4 - Incerteza - Segunda Leitura': string;
  'V.I.I 1.1 - Segunda Leitura': string;
  'V.I.I 1.2 - Segunda Leitura': string;
  'V.I.I 1.3 - Segunda Leitura': string;
  'V.I.I 2.1 - Segunda Leitura': string;
  'V.I.I 2.2 - Segunda Leitura': string;
  'V.I.I 2.3 - Segunda Leitura': string;
  'V.I.I 3.1 - Segunda Leitura': string;
  'V.I.I 3.2 - Segunda Leitura': string;
  'V.I.I 3.3 - Segunda Leitura': string;
  'V.I.I 4.1 - Segunda Leitura': string;
  'V.I.I 4.2 - Segunda Leitura': string;
  'V.I.I 4.3 - Segunda Leitura': string;

  'V.C.C 1 - Terceira Leitura': string;
  'V.C.C 1 - Incerteza - Terceira Leitura': string;
  'V.C.C 2 - Terceira Leitura': string;
  'V.C.C 2 - Incerteza - Terceira Leitura': string;
  'V.C.C 3 - Terceira Leitura': string;
  'V.C.C 3 - Incerteza - Terceira Leitura': string;
  'V.C.C 4 - Terceira Leitura': string;
  'V.C.C 4 - Incerteza - Terceira Leitura': string;
  'V.I.I 1.1 - Terceira Leitura': string;
  'V.I.I 1.2 - Terceira Leitura': string;
  'V.I.I 1.3 - Terceira Leitura': string;
  'V.I.I 2.1 - Terceira Leitura': string;
  'V.I.I 2.2 - Terceira Leitura': string;
  'V.I.I 2.3 - Terceira Leitura': string;
  'V.I.I 3.1 - Terceira Leitura': string;
  'V.I.I 3.2 - Terceira Leitura': string;
  'V.I.I 3.3 - Terceira Leitura': string;
  'V.I.I 4.1 - Terceira Leitura': string;
  'V.I.I 4.2 - Terceira Leitura': string;
  'V.I.I 4.3 - Terceira Leitura': string;

  // Dados Adicionais
  'Padrão': string;
  'Local da Calibração':string;
  'Temperatura do Local da Calibração': string;
  'Umidade Relativa do Local da Calibração': string;
  'Próxima Data de Calibração': string;
}

@Component({
  selector: 'explorer-body',
  templateUrl: './explorer.body.html',
  styleUrls: ['./explorer.body.css']
})

export class ExplorerBodyComponent implements OnInit {

  certCollection: AngularFirestoreCollection<Certificado>;
  cert: Observable<Certificado[]>;
  data: any;
  dataSource;

  constructor(private afs: AngularFirestore) { }

  ngOnInit() {
    this.certCollection = this.afs.collection('preciso-certificados')
    this.cert = this.certCollection.valueChanges()
    this.data = this.cert.subscribe(certificados => { this.dataSource = new MatTableDataSource(certificados);})
  }

  displayedColumns: string[] = ['Select', 'Empresa', 'Tipo de Instrumento', 'Instrumento', 'Marca do Instrumento', 'Modelo do Instrumento'];
  selection = new SelectionModel<Certificado>(true, [])

  isAllSelected() {
    const numSelected = this.selection.selected.length;
    const numRows = this.dataSource.data.length;
    return numSelected === numRows;
  }

  masterToggle() {
    this.isAllSelected() ?
      this.selection.clear() :
      this.dataSource.data.forEach(row => this.selection.select(row));
  }

  applyFilter(filterValue: string) {
    this.dataSource.filter = filterValue.trim().toLowerCase();
  }

  exportData(selectedData: any) {
    var selectedCert = this.selection.selected;
    var filepath = "";
    console.log(selectedCert.map( data => data.Empresa));
    var xmlhttp = new XMLHttpRequest();

    switch (selectedCert.map((data => data["Tipo de Instrumento"])).toString()){
      case "Medidor de Pressão":
        filepath = "../src/app/preciso-explorer/templates/preciso-medidor-pressao.xlsx";
      break;
      case "Termohigrometro":
        filepath = "../src/app/preciso-explorer/templates/default.xlsx";
      break;
      case "Vidraria Graduada":
        filepath = "../src/app/preciso-explorer/templates/preciso-vidraria-graduada.xlsx";
      break;
    }

    xmlhttp.open("GET", filepath, true);
    xmlhttp.responseType = "arraybuffer";
    xmlhttp.onreadystatechange = function () {
      if (xmlhttp.readyState === 4) {
        if (xmlhttp.status === 200) {

          XLSX.fromDataAsync(xmlhttp.response).then(workbook => {
            if (selectedCert != null){
              
              switch (selectedCert.map((data => data["Tipo de Instrumento"])).toString()){
                case "Medidor de Pressão":
                  workbook.sheet(0).cell('AV5').value((selectedCert.map(data => data.Empresa)).toString());
                  workbook.sheet(0).cell('AV6').value((selectedCert.map(data => data.Endereço)).toString());
                  workbook.sheet(0).cell('AV7').value((selectedCert.map(data => data["Cidade c/ Estado"])).toString());
                  workbook.sheet(0).cell('AV9').value((selectedCert.map(data => data["Tipo de Instrumento"])).toString());
                  workbook.sheet(0).cell('AV11').value((selectedCert.map(data => data.Equipamento)).toString());
                  workbook.sheet(0).cell('AV8').value((selectedCert.map(data => data.Instrumento)).toString());
                  workbook.sheet(0).cell('AV12').value((selectedCert.map(data => data["Marca do Instrumento"])).toString());
                  workbook.sheet(0).cell('AV13').value((selectedCert.map(data => data["Modelo do Instrumento"])).toString());
                  workbook.sheet(0).cell('AV10').value((selectedCert.map(data => data["Classe do Instrumento"])).toString());
                  workbook.sheet(0).cell('AV14').value((selectedCert.map(data => data["Número de Série do Instrumento"])).toString());
                  workbook.sheet(0).cell('AV15').value((selectedCert.map(data => data["Identificação do Instrumento"])).toString());
                  workbook.sheet(0).cell('AV17').value((selectedCert.map(data => data["Inicio da Faixa de Uso"])).toString());
                  workbook.sheet(0).cell('AX17').value((selectedCert.map(data => data["Final da Faixa de Uso"])).toString());
                  workbook.sheet(0).cell('AY17').value((selectedCert.map(data => data.Grandeza)).toString());
                  workbook.sheet(0).cell('AV18').value((selectedCert.map(data => data["Valor de uma Divisão"])).toString());
                break;
                case "Termohigrometro":
                  workbook.sheet(0).cell('AW5').value((selectedCert.map(data => data.Empresa)).toString());
                  workbook.sheet(0).cell('AW6').value((selectedCert.map(data => data.Endereço)).toString());
                  workbook.sheet(0).cell('AW7').value((selectedCert.map(data => data["Cidade c/ Estado"])).toString());
                  workbook.sheet(0).cell('AW8').value((selectedCert.map(data => data["Tipo de Instrumento"])).toString());
                  workbook.sheet(0).cell('AW9').value((selectedCert.map(data => data.Equipamento)).toString());
                  workbook.sheet(0).cell('AW10').value((selectedCert.map(data => data["Marca do Instrumento"])).toString());
                  workbook.sheet(0).cell('AW11').value((selectedCert.map(data => data["Modelo do Instrumento"])).toString());
                  workbook.sheet(0).cell('AW12').value((selectedCert.map(data => data["Classe do Instrumento"])).toString());
                  workbook.sheet(0).cell('AW13').value((selectedCert.map(data => data["Número de Série do Instrumento"])).toString());
                  workbook.sheet(0).cell('AW14').value((selectedCert.map(data => data["Identificação do Instrumento"])).toString());
                break;
                case "Vidraria Graduada":
                  workbook.sheet(0).cell('AV5').value((selectedCert.map(data => data.Empresa)).toString());
                  workbook.sheet(0).cell('AV6').value((selectedCert.map(data => data.Endereço)).toString());
                  workbook.sheet(0).cell('AV7').value((selectedCert.map(data => data["Cidade c/ Estado"])).toString());
                  workbook.sheet(0).cell('AV8').value((selectedCert.map(data => data["Tipo de Instrumento"])).toString());
                  workbook.sheet(0).cell('AV9').value((selectedCert.map(data => data.Equipamento)).toString());
                  workbook.sheet(0).cell('AV10').value((selectedCert.map(data => data["Marca do Instrumento"])).toString());
                  workbook.sheet(0).cell('AV11').value((selectedCert.map(data => data["Modelo do Instrumento"])).toString());
                  workbook.sheet(0).cell('AV12').value((selectedCert.map(data => data["Classe do Instrumento"])).toString());
                  workbook.sheet(0).cell('AV13').value((selectedCert.map(data => data["Número de Série do Instrumento"])).toString());
                  workbook.sheet(0).cell('AV14').value((selectedCert.map(data => data["Identificação do Instrumento"])).toString());
                break;
              }
            }
            workbook.outputAsync().then(function (blob) {
             if (window.navigator && window.navigator.msSaveOrOpenBlob) {
                // If IE, you must uses a different method.
                window.navigator.msSaveOrOpenBlob(blob, "out.xlsx");
              } else {
                var url = window.URL.createObjectURL(blob);
                var a = document.createElement("a");
                document.body.appendChild(a);
                a.href = url;
                a.download = "out.xlsx";
                a.click();
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);
              }
            })
          }
          )
        }
      }
    }
    xmlhttp.send();
  }
}



