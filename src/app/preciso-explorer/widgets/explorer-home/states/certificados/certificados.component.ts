import { Component, OnInit } from '@angular/core';
import { Observable } from 'rxjs';
import 'rxjs/add/operator/map';

import { AngularFirestore, AngularFirestoreCollection } from 'angularfire2/firestore';
import { MatTableDataSource } from '@angular/material';
import { SelectionModel } from '@angular/cdk/collections';
import * as XLSX from 'xlsx-populate';

interface Certificado {
  PrecisoID: string;
  ID: string;
  Ano: string;
  'Mês': string;
  Incremental: string;
  'Hora do Envio':string;
  'IDTipo':string;

  Empresa: string;
  'Endereço': string;
  'Cidade c/ Estado': string;
  Equipamento: string;
  'Tipo de Instrumento': string;
  Instrumento: string;

  Unidade: string;
  'Inicio da Faixa de Uso': string;
  'Final da Faixa de Uso': string;
  'Valor de uma Divisao': string;
  'Inicio de Escala': string;
  'Final de Escala': string;
  'Volume Nominal': string;
  'Pressao Atmosferica': string;

  'Marca do Instrumento': string;
  'Modelo do Instrumento': string;
  'Classe do Instrumento': string;
  'Classe': string;
  'Número de Série do Instrumento': string;
  'Identificação do Instrumento': string;

  'Temperatura de Entrada: Inicio de Escala':string;
  'Temperatura de Entrada: Final de Escala': string;
  'Temperatura de Entrada: Inicio da Faixa de Uso': string;
  'Temperatura de Entrada: Final da Faixa de Uso': string;
  'Temperatura de Entrada: Valor de uma Divisão': string;
  'Temperatura de Saida: Inicio de Escala': string;
  'Temperatura de Saida: Final de Escala': string;
  'Temperatura de Saida: Inicio da Faixa de Uso': string;
  'Temperatura de Saida: Final da Faixa de Uso': string;
  'Temperatura de Saida: Valor de uma Divisão': string;
  'Umidade Relativa: Inicio de Escala': string;
  'Umidade Relativa: Final de Escala': string;
  'Umidade Relativa: Inicio da Faixa de Uso': string;
  'Umidade Relativa: Final da Faixa de Uso': string;
  'Umidade Relativa: Valor de uma Divisão': string;

  // -------------------------------------------------------------------------------------------------------------------------------- //
  'V.V.C 1 - Leitura 1': string;
  'V.V.C 2 - Leitura 1': string;
  'V.V.C 3 - Leitura 1': string;
  'V.V.C 4 - Leitura 1': string;
  'V.V.C 5 - Leitura 1': string;
  'V.I.I 1 - Leitura 1 - Linha 1': string;
  'V.I.I 2 - Leitura 1 - Linha 1': string;
  'V.I.I 3 - Leitura 1 - Linha 1': string;
  'V.I.I 1 - Leitura 1 - Linha 2': string;
  'V.I.I 2 - Leitura 1 - Linha 2': string;
  'V.I.I 3 - Leitura 1 - Linha 2': string;
  'V.I.I 1 - Leitura 1 - Linha 3': string;
  'V.I.I 2 - Leitura 1 - Linha 3': string;
  'V.I.I 3 - Leitura 1 - Linha 3': string;
  'V.I.I 1 - Leitura 1 - Linha 4': string;
  'V.I.I 2 - Leitura 1 - Linha 4': string;
  'V.I.I 3 - Leitura 1 - Linha 4': string;
  'V.I.I 1 - Leitura 1 - Linha 5': string;
  'V.I.I 2 - Leitura 1 - Linha 5': string;
  'V.I.I 3 - Leitura 1 - Linha 5': string;

  'V.V.C 1 - Leitura 2': string;
  'V.V.C 2 - Leitura 2': string;
  'V.V.C 3 - Leitura 2': string;
  'V.V.C 4 - Leitura 2': string;
  'V.V.C 5 - Leitura 2': string;
  'V.I.I 1 - Leitura 2 - Linha 1': string;
  'V.I.I 2 - Leitura 2 - Linha 1': string;
  'V.I.I 3 - Leitura 2 - Linha 1': string;
  'V.I.I 1 - Leitura 2 - Linha 2': string;
  'V.I.I 2 - Leitura 2 - Linha 2': string;
  'V.I.I 3 - Leitura 2 - Linha 2': string;
  'V.I.I 1 - Leitura 2 - Linha 3': string;
  'V.I.I 2 - Leitura 2 - Linha 3': string;
  'V.I.I 3 - Leitura 2 - Linha 3': string;
  'V.I.I 1 - Leitura 2 - Linha 4': string;
  'V.I.I 2 - Leitura 2 - Linha 4': string;
  'V.I.I 3 - Leitura 2 - Linha 4': string;
  'V.I.I 1 - Leitura 2 - Linha 5': string;
  'V.I.I 2 - Leitura 2 - Linha 5': string;
  'V.I.I 3 - Leitura 2 - Linha 5': string;

  'V.V.C 1 - Leitura 3': string;
  'V.V.C 2 - Leitura 3': string;
  'V.V.C 3 - Leitura 3': string;
  'V.V.C 4 - Leitura 3': string;
  'V.V.C 5 - Leitura 3': string;
  'V.I.I 1 - Leitura 3 - Linha 1': string;
  'V.I.I 2 - Leitura 3 - Linha 1': string;
  'V.I.I 3 - Leitura 3 - Linha 1': string;
  'V.I.I 1 - Leitura 3 - Linha 2': string;
  'V.I.I 2 - Leitura 3 - Linha 2': string;
  'V.I.I 3 - Leitura 3 - Linha 2': string;
  'V.I.I 1 - Leitura 3 - Linha 3': string;
  'V.I.I 2 - Leitura 3 - Linha 3': string;
  'V.I.I 3 - Leitura 3 - Linha 3': string;
  'V.I.I 1 - Leitura 3 - Linha 4': string;
  'V.I.I 2 - Leitura 3 - Linha 4': string;
  'V.I.I 3 - Leitura 3 - Linha 4': string;
  'V.I.I 1 - Leitura 3 - Linha 5': string;
  'V.I.I 2 - Leitura 3 - Linha 5': string;
  'V.I.I 3 - Leitura 3 - Linha 5': string;

  // ------------------------------------------------------------------------------------------------------------------------- //
  'V.I.I 1 - Leitura 1': string;
  'V.I.I 2 - Leitura 1': string;
  'V.I.I 3 - Leitura 1': string;
  'V.I.I 4 - Leitura 1': string;
  'V.I.I 5 - Leitura 1': string;
  'V.I.I 6 - Leitura 1': string;
  'V.I.I 7 - Leitura 1': string;
  'V.I.I 8 - Leitura 1': string;
  'V.I.I 9 - Leitura 1': string;
  'V.I.I 10 - Leitura 1': string;
  'V.V.C 1 - Leitura 1 - Linha 1':string;
  'V.V.C 2 - Leitura 1 - Linha 1': string;
  'V.V.C 3 - Leitura 1 - Linha 1': string;
  'V.V.C 1 - Leitura 1 - Linha 2': string;
  'V.V.C 2 - Leitura 1 - Linha 2': string;
  'V.V.C 3 - Leitura 1 - Linha 2': string;
  'V.V.C 1 - Leitura 1 - Linha 3': string;
  'V.V.C 2 - Leitura 1 - Linha 3': string;
  'V.V.C 3 - Leitura 1 - Linha 3': string;
  'V.V.C 1 - Leitura 1 - Linha 4': string;
  'V.V.C 2 - Leitura 1 - Linha 4': string;
  'V.V.C 3 - Leitura 1 - Linha 4': string;
  'V.V.C 1 - Leitura 1 - Linha 5': string;
  'V.V.C 2 - Leitura 1 - Linha 5': string;
  'V.V.C 3 - Leitura 1 - Linha 5': string;
  'V.V.C 1 - Leitura 1 - Linha 6': string;
  'V.V.C 2 - Leitura 1 - Linha 6': string;
  'V.V.C 3 - Leitura 1 - Linha 6': string;
  'V.V.C 1 - Leitura 1 - Linha 7': string;
  'V.V.C 2 - Leitura 1 - Linha 7': string;
  'V.V.C 3 - Leitura 1 - Linha 7': string;
  'V.V.C 1 - Leitura 1 - Linha 8': string;
  'V.V.C 2 - Leitura 1 - Linha 8': string;
  'V.V.C 3 - Leitura 1 - Linha 8': string;
  'V.V.C 1 - Leitura 1 - Linha 9': string;
  'V.V.C 2 - Leitura 1 - Linha 9': string;
  'V.V.C 3 - Leitura 1 - Linha 9': string;
  'V.V.C 1 - Leitura 1 - Linha 10': string;
  'V.V.C 2 - Leitura 1 - Linha 10': string;
  'V.V.C 3 - Leitura 1 - Linha 10': string;

  'V.I.I 1 - Leitura 2': string;
  'V.I.I 2 - Leitura 2': string;
  'V.I.I 3 - Leitura 2': string;
  'V.I.I 4 - Leitura 2': string;
  'V.I.I 5 - Leitura 2': string;
  'V.I.I 6 - Leitura 2': string;
  'V.I.I 7 - Leitura 2': string;
  'V.I.I 8 - Leitura 2': string;
  'V.I.I 9 - Leitura 2': string;
  'V.I.I 10 - Leitura 2': string;
  'V.V.C 1 - Leitura 2 - Linha 1': string;
  'V.V.C 2 - Leitura 2 - Linha 1': string;
  'V.V.C 3 - Leitura 2 - Linha 1': string;
  'V.V.C 1 - Leitura 2 - Linha 2': string;
  'V.V.C 2 - Leitura 2 - Linha 2': string;
  'V.V.C 3 - Leitura 2 - Linha 2': string;
  'V.V.C 1 - Leitura 2 - Linha 3': string;
  'V.V.C 2 - Leitura 2 - Linha 3': string;
  'V.V.C 3 - Leitura 2 - Linha 3': string;
  'V.V.C 1 - Leitura 2 - Linha 4': string;
  'V.V.C 2 - Leitura 2 - Linha 4': string;
  'V.V.C 3 - Leitura 2 - Linha 4': string;
  'V.V.C 1 - Leitura 2 - Linha 5': string;
  'V.V.C 2 - Leitura 2 - Linha 5': string;
  'V.V.C 3 - Leitura 2 - Linha 5': string;
  'V.V.C 1 - Leitura 2 - Linha 6': string;
  'V.V.C 2 - Leitura 2 - Linha 6': string;
  'V.V.C 3 - Leitura 2 - Linha 6': string;
  'V.V.C 1 - Leitura 2 - Linha 7': string;
  'V.V.C 2 - Leitura 2 - Linha 7': string;
  'V.V.C 3 - Leitura 2 - Linha 7': string;
  'V.V.C 1 - Leitura 2 - Linha 8': string;
  'V.V.C 2 - Leitura 2 - Linha 8': string;
  'V.V.C 3 - Leitura 2 - Linha 8': string;
  'V.V.C 1 - Leitura 2 - Linha 9': string;
  'V.V.C 2 - Leitura 2 - Linha 9': string;
  'V.V.C 3 - Leitura 2 - Linha 9': string;
  'V.V.C 1 - Leitura 2 - Linha 10': string;
  'V.V.C 2 - Leitura 2 - Linha 10': string;
  'V.V.C 3 - Leitura 2 - Linha 10': string;

  'Padrão 1': string;
  'Padrão 2': string;
  'Padrão 3': string;

  'Data de Calibração': string;
  'Local da Calibração': string;
  'Temperatura do Local da Calibração': string;
  'Umidade Relativa do Local da Calibração': string;
  'Próxima Data de Calibração': string;
}

@Component({
  selector: 'app-certificados',
  templateUrl: './certificados.component.html',
  styleUrls: ['./certificados.component.css']
})

export class CertificadosComponent implements OnInit {

  certCollection: AngularFirestoreCollection<Certificado>;
  cert: Observable<Certificado[]>;
  data: any;
  dataSource;

  constructor(private afs: AngularFirestore) { }

  ngOnInit() {
    this.certCollection = this.afs.collection('preciso-certificados')
    this.cert = this.certCollection.valueChanges()
    this.data = this.cert.subscribe(certificados => { this.dataSource = new MatTableDataSource(certificados); })
  }

  displayedColumns: string[] = ['Select', 'Empresa', 'Tipo de Instrumento', 'Instrumento', 'Marca do Instrumento', 'Modelo do Instrumento'];
  selection = new SelectionModel<Certificado>(true, []);

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

  exportData() {
    var selectedCert = this.selection.selected;
    var filepath = "";
    var xmlhttp = new XMLHttpRequest();
   
    switch (selectedCert.map((data => data.IDTipo)).toString()) {
        case "R101180813F":
          filepath = "https://github.com/dev-whiteflag/preciso-metrologia-webapp/raw/master/src/app/preciso-explorer/templates/preciso-medidor-pressao.xlsx";
          break;
        case "R05150707I-3":
          filepath = "https://github.com/dev-whiteflag/preciso-metrologia-webapp/raw/master/src/app/preciso-explorer/templates/preciso-termohigrometro3.xlsx";
          break;
        case "R100180915F":
          filepath = "https://github.com/dev-whiteflag/preciso-metrologia-webapp/raw/master/src/app/preciso-explorer/templates/preciso-vidraria-graduada.xlsx";
          break;
        case "R101181123A":
          filepath = "https://github.com/dev-whiteflag/preciso-metrologia-webapp/raw/master/src/app/preciso-explorer/templates/preciso-medidor-pressao10.xlsx";
          break;
      }
      xmlhttp.open("GET", filepath, true);
      xmlhttp.responseType = "arraybuffer";
      xmlhttp.onreadystatechange = function () {
        console.log("xmlHttp: stateChange ok");
        if (xmlhttp.readyState === 4) {
          console.log("xmlHttp: State 4 ok");
          if (xmlhttp.status === 200) {
            console.log("xmlHttp: Status 200 ok");
          
            XLSX.fromDataAsync(xmlhttp.response).then(workbook => {
              console.log("dataAsync ok");
              if (selectedCert != null) {
                var id = (selectedCert.map(data => data["ID"]).toString());
                switch (selectedCert.map((data => data.IDTipo)).toString()) {
                  case "R101180813F":
                    workbook.sheet(0).cell('AV3').value((selectedCert.map(data => data.Ano)).toString());
                    workbook.sheet(0).cell('AW3').value((selectedCert.map(data => data.Mês)).toString());
                    workbook.sheet(0).cell('AX3').value((selectedCert.map(data => data.PrecisoID)).toString());
                    workbook.sheet(0).cell('AV4').value((selectedCert.map(data => data.Incremental)).toString());

                    workbook.sheet(0).cell('AV5').value((selectedCert.map(data => data.Empresa)).toString());
                    workbook.sheet(0).cell('AV6').value((selectedCert.map(data => data.Endereço)).toString());
                    workbook.sheet(0).cell('AV7').value((selectedCert.map(data => data["Cidade c/ Estado"])).toString());
                    workbook.sheet(0).cell('AV8').value((selectedCert.map(data => data.Equipamento)).toString());
                    workbook.sheet(0).cell('AV9').value((selectedCert.map(data => data.Instrumento)).toString());
                    workbook.sheet(0).cell('AV10').value((selectedCert.map(data => data.Classe)).toString());
                    workbook.sheet(0).cell('AV11').value((selectedCert.map(data => data.Equipamento)).toString());

                    workbook.sheet(0).cell('AV12').value((selectedCert.map(data => data["Marca do Instrumento"])).toString());
                    workbook.sheet(0).cell('AV13').value((selectedCert.map(data => data["Modelo do Instrumento"])).toString());
                    workbook.sheet(0).cell('AV14').value((selectedCert.map(data => data["Número de Série do Instrumento"])).toString());
                    workbook.sheet(0).cell('AV15').value((selectedCert.map(data => data["Identificação do Instrumento"])).toString());

                    workbook.sheet(0).cell('AV16').value(parseFloat(selectedCert.map(data => (data["Inicio de Escala"])).toString()));
                    workbook.sheet(0).cell('AX16').value(parseFloat(selectedCert.map(data => (data["Final de Escala"])).toString()));

                    workbook.sheet(0).cell('AV17').value(parseFloat(selectedCert.map(data => (data["Inicio da Faixa de Uso"])).toString()));
                    workbook.sheet(0).cell('AX17').value(parseFloat(selectedCert.map(data => (data["Final da Faixa de Uso"])).toString()));
                    workbook.sheet(0).cell('AY17').value((selectedCert.map(data => data.Unidade)).toString());
                    workbook.sheet(0).cell('AV18').value(parseFloat(selectedCert.map(data => (data["Valor de uma Divisao"])).toString()));

                    workbook.sheet(0).cell('AT23').value(parseFloat(selectedCert.map(data => (data["V.I.I 1 - Leitura 1"])).toString()));
                    workbook.sheet(0).cell('AT24').value(parseFloat(selectedCert.map(data => (data["V.I.I 2 - Leitura 1"])).toString()));
                    workbook.sheet(0).cell('AT25').value(parseFloat(selectedCert.map(data => (data["V.I.I 3 - Leitura 1"])).toString()));
                    workbook.sheet(0).cell('AT26').value(parseFloat(selectedCert.map(data => (data["V.I.I 4 - Leitura 1"])).toString()));
                    workbook.sheet(0).cell('AT27').value(parseFloat(selectedCert.map(data => (data["V.I.I 5 - Leitura 1"])).toString()));

                    workbook.sheet(0).cell('AU23').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 1 - Linha 1"])).toString()));
                    workbook.sheet(0).cell('AV23').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 1 - Linha 1"])).toString()));
                    workbook.sheet(0).cell('AW23').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 1 - Linha 1"])).toString()));
                    workbook.sheet(0).cell('AU24').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 1 - Linha 2"])).toString()));
                    workbook.sheet(0).cell('AV24').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 1 - Linha 2"])).toString()));
                    workbook.sheet(0).cell('AW24').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 1 - Linha 2"])).toString()));
                    workbook.sheet(0).cell('AU25').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 1 - Linha 3"])).toString()));
                    workbook.sheet(0).cell('AV25').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 1 - Linha 3"])).toString()));
                    workbook.sheet(0).cell('AW25').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 1 - Linha 3"])).toString()));
                    workbook.sheet(0).cell('AU26').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 1 - Linha 4"])).toString()));
                    workbook.sheet(0).cell('AV26').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 1 - Linha 4"])).toString()));
                    workbook.sheet(0).cell('AW26').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 1 - Linha 4"])).toString()));
                    workbook.sheet(0).cell('AU27').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 1 - Linha 5"])).toString()));
                    workbook.sheet(0).cell('AV27').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 1 - Linha 5"])).toString()));
                    workbook.sheet(0).cell('AW27').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 1 - Linha 5"])).toString()));

                    workbook.sheet(0).cell('AU31').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 2 - Linha 1"])).toString()));
                    workbook.sheet(0).cell('AV31').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 2 - Linha 1"])).toString()));
                    workbook.sheet(0).cell('AW31').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 2 - Linha 1"])).toString()));
                    workbook.sheet(0).cell('AU32').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 2 - Linha 2"])).toString()));
                    workbook.sheet(0).cell('AV32').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 2 - Linha 2"])).toString()));
                    workbook.sheet(0).cell('AW32').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 2 - Linha 2"])).toString()));
                    workbook.sheet(0).cell('AU33').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 2 - Linha 3"])).toString()));
                    workbook.sheet(0).cell('AV33').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 2 - Linha 3"])).toString()));
                    workbook.sheet(0).cell('AW33').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 2 - Linha 3"])).toString()));
                    workbook.sheet(0).cell('AU34').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 2 - Linha 4"])).toString()));
                    workbook.sheet(0).cell('AV34').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 2 - Linha 4"])).toString()));
                    workbook.sheet(0).cell('AW34').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 2 - Linha 4"])).toString()));
                    workbook.sheet(0).cell('AU35').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 2 - Linha 5"])).toString()));
                    workbook.sheet(0).cell('AV35').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 2 - Linha 5"])).toString()));
                    workbook.sheet(0).cell('AW35').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 2 - Linha 5"])).toString()));

                    workbook.sheet(0).cell('AV39').value((selectedCert.map(data => data["Local da Calibração"])).toString());
                    workbook.sheet(0).cell('AV40').value(parseFloat(selectedCert.map(data => (data["Temperatura do Local da Calibração"])).toString()));
                    workbook.sheet(0).cell('AV41').value(parseFloat(selectedCert.map(data => (data["Umidade Relativa do Local da Calibração"])).toString()));

                    workbook.sheet(0).cell('AT44').value((selectedCert.map(data => data["Padrão 1"])).toString());
                    workbook.sheet(0).cell('AV46').value((selectedCert.map(data => data["Data de Calibração"])).toString());
                    workbook.sheet(0).cell('AV47').value((selectedCert.map(data => data["Próxima Data de Calibração"])).toString());
                    workbook.sheet(0).cell('AV48').value("François Fernandez");
                    break;
                  
                    case 'R101181123A':
                    console.log(selectedCert);
                    workbook.sheet(0).cell('AV3').value((selectedCert.map(data => data.Ano)).toString());
                    workbook.sheet(0).cell('AW3').value((selectedCert.map(data => data.Mês)).toString());
                    workbook.sheet(0).cell('AX3').value((selectedCert.map(data => data.PrecisoID)).toString());
                    workbook.sheet(0).cell('AV4').value((selectedCert.map(data => data.Incremental)).toString());

                    workbook.sheet(0).cell('AV5').value((selectedCert.map(data => data.Empresa)).toString());
                    workbook.sheet(0).cell('AV6').value((selectedCert.map(data => data.Endereço)).toString());
                    workbook.sheet(0).cell('AV7').value((selectedCert.map(data => data["Cidade c/ Estado"])).toString());
                    workbook.sheet(0).cell('AV8').value((selectedCert.map(data => data.Instrumento)).toString());
                    workbook.sheet(0).cell('AV9').value((selectedCert.map(data => data.Classe)).toString());
                    workbook.sheet(0).cell('AV10').value((selectedCert.map(data => data.Equipamento)).toString());

                    workbook.sheet(0).cell('AV11').value((selectedCert.map(data => data["Marca do Instrumento"])).toString());
                    workbook.sheet(0).cell('AV12').value((selectedCert.map(data => data["Modelo do Instrumento"])).toString());
                    workbook.sheet(0).cell('AV13').value((selectedCert.map(data => data["Classe do Instrumento"])).toString());
                    workbook.sheet(0).cell('AV14').value((selectedCert.map(data => data["Número de Série do Instrumento"])).toString());
                    workbook.sheet(0).cell('AV15').value((selectedCert.map(data => data["Identificação do Instrumento"])).toString());

                    workbook.sheet(0).cell('AV16').value(parseFloat(selectedCert.map(data => (data["Inicio de Escala"])).toString()));
                    workbook.sheet(0).cell('AX16').value(parseFloat(selectedCert.map(data => (data["Final de Escala"])).toString()));

                    workbook.sheet(0).cell('AV17').value(parseFloat(selectedCert.map(data => (data["Inicio da Faixa de Uso"])).toString()));
                    workbook.sheet(0).cell('AX17').value(parseFloat(selectedCert.map(data => (data["Final da Faixa de Uso"])).toString()));
                    workbook.sheet(0).cell('AY16').value((selectedCert.map(data => data.Unidade)).toString());
                    workbook.sheet(0).cell('AV18').value(parseFloat(selectedCert.map(data => (data["Valor de uma Divisao"])).toString()));

                    workbook.sheet(0).cell('AT22').value(parseFloat(selectedCert.map(data => (data["V.I.I 1 - Leitura 1"])).toString()));
                    workbook.sheet(0).cell('AT23').value(parseFloat(selectedCert.map(data => (data["V.I.I 2 - Leitura 1"])).toString()));
                    workbook.sheet(0).cell('AT24').value(parseFloat(selectedCert.map(data => (data["V.I.I 3 - Leitura 1"])).toString()));
                    workbook.sheet(0).cell('AT25').value(parseFloat(selectedCert.map(data => (data["V.I.I 4 - Leitura 1"])).toString()));
                    workbook.sheet(0).cell('AT26').value(parseFloat(selectedCert.map(data => (data["V.I.I 5 - Leitura 1"])).toString()));
                    workbook.sheet(0).cell('AT27').value(parseFloat(selectedCert.map(data => (data["V.I.I 6 - Leitura 1"])).toString()));
                    workbook.sheet(0).cell('AT28').value(parseFloat(selectedCert.map(data => (data["V.I.I 7 - Leitura 1"])).toString()));
                    workbook.sheet(0).cell('AT29').value(parseFloat(selectedCert.map(data => (data["V.I.I 8 - Leitura 1"])).toString()));
                    workbook.sheet(0).cell('AT30').value(parseFloat(selectedCert.map(data => (data["V.I.I 9 - Leitura 1"])).toString()));
                    workbook.sheet(0).cell('AT31').value(parseFloat(selectedCert.map(data => (data["V.I.I 10 - Leitura 1"])).toString()));

                    workbook.sheet(0).cell('AU22').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 1 - Linha 1"])).toString()));
                    workbook.sheet(0).cell('AV22').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 1 - Linha 1"])).toString()));
                    workbook.sheet(0).cell('AW22').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 1 - Linha 1"])).toString()));
                    workbook.sheet(0).cell('AU23').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 1 - Linha 2"])).toString()));
                    workbook.sheet(0).cell('AV23').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 1 - Linha 2"])).toString()));
                    workbook.sheet(0).cell('AW23').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 1 - Linha 2"])).toString()));
                    workbook.sheet(0).cell('AU24').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 1 - Linha 3"])).toString()));
                    workbook.sheet(0).cell('AV24').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 1 - Linha 3"])).toString()));
                    workbook.sheet(0).cell('AW24').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 1 - Linha 3"])).toString()));
                    workbook.sheet(0).cell('AU25').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 1 - Linha 4"])).toString()));
                    workbook.sheet(0).cell('AV25').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 1 - Linha 4"])).toString()));
                    workbook.sheet(0).cell('AW25').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 1 - Linha 4"])).toString()));
                    workbook.sheet(0).cell('AU26').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 1 - Linha 5"])).toString()));
                    workbook.sheet(0).cell('AV26').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 1 - Linha 5"])).toString()));
                    workbook.sheet(0).cell('AW26').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 1 - Linha 5"])).toString()));
                    workbook.sheet(0).cell('AU27').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 1 - Linha 6"])).toString()));
                    workbook.sheet(0).cell('AV27').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 1 - Linha 6"])).toString()));
                    workbook.sheet(0).cell('AW27').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 1 - Linha 6"])).toString()));
                    workbook.sheet(0).cell('AU28').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 1 - Linha 7"])).toString()));
                    workbook.sheet(0).cell('AV28').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 1 - Linha 7"])).toString()));
                    workbook.sheet(0).cell('AW28').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 1 - Linha 7"])).toString()));
                    workbook.sheet(0).cell('AU29').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 1 - Linha 8"])).toString()));
                    workbook.sheet(0).cell('AV29').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 1 - Linha 8"])).toString()));
                    workbook.sheet(0).cell('AW29').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 1 - Linha 8"])).toString()));
                    workbook.sheet(0).cell('AU30').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 1 - Linha 9"])).toString()));
                    workbook.sheet(0).cell('AV30').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 1 - Linha 9"])).toString()));
                    workbook.sheet(0).cell('AW30').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 1 - Linha 9"])).toString()));
                    workbook.sheet(0).cell('AU31').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 1 - Linha 10"])).toString()));
                    workbook.sheet(0).cell('AV31').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 1 - Linha 10"])).toString()));
                    workbook.sheet(0).cell('AW31').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 1 - Linha 10"])).toString()));

                    workbook.sheet(0).cell('AU36').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 2 - Linha 1"])).toString()));
                    workbook.sheet(0).cell('AV36').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 2 - Linha 1"])).toString()));
                    workbook.sheet(0).cell('AW36').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 2 - Linha 1"])).toString()));
                    workbook.sheet(0).cell('AU37').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 2 - Linha 2"])).toString()));
                    workbook.sheet(0).cell('AV37').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 2 - Linha 2"])).toString()));
                    workbook.sheet(0).cell('AW37').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 2 - Linha 2"])).toString()));
                    workbook.sheet(0).cell('AU38').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 2 - Linha 3"])).toString()));
                    workbook.sheet(0).cell('AV38').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 2 - Linha 3"])).toString()));
                    workbook.sheet(0).cell('AW38').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 2 - Linha 3"])).toString()));
                    workbook.sheet(0).cell('AU39').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 2 - Linha 4"])).toString()));
                    workbook.sheet(0).cell('AV39').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 2 - Linha 4"])).toString()));
                    workbook.sheet(0).cell('AW39').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 2 - Linha 4"])).toString()));
                    workbook.sheet(0).cell('AU40').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 2 - Linha 5"])).toString()));
                    workbook.sheet(0).cell('AV40').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 2 - Linha 5"])).toString()));
                    workbook.sheet(0).cell('AW40').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 2 - Linha 5"])).toString()));
                    workbook.sheet(0).cell('AU41').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 2 - Linha 6"])).toString()));
                    workbook.sheet(0).cell('AV41').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 2 - Linha 6"])).toString()));
                    workbook.sheet(0).cell('AW41').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 2 - Linha 6"])).toString()));
                    workbook.sheet(0).cell('AU42').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 2 - Linha 7"])).toString()));
                    workbook.sheet(0).cell('AV42').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 2 - Linha 7"])).toString()));
                    workbook.sheet(0).cell('AW42').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 2 - Linha 7"])).toString()));
                    workbook.sheet(0).cell('AU43').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 2 - Linha 8"])).toString()));
                    workbook.sheet(0).cell('AV43').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 2 - Linha 8"])).toString()));
                    workbook.sheet(0).cell('AW43').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 2 - Linha 8"])).toString()));
                    workbook.sheet(0).cell('AU44').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 2 - Linha 9"])).toString()));
                    workbook.sheet(0).cell('AV44').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 2 - Linha 9"])).toString()));
                    workbook.sheet(0).cell('AW44').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 2 - Linha 9"])).toString()));
                    workbook.sheet(0).cell('AU45').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 2 - Linha 10"])).toString()));
                    workbook.sheet(0).cell('AV45').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 2 - Linha 10"])).toString()));
                    workbook.sheet(0).cell('AW45').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 2 - Linha 10"])).toString()));
                    workbook.sheet(0).cell('AV49').value((selectedCert.map(data => data["Local da Calibração"])).toString());
                    workbook.sheet(0).cell('AV50').value(parseFloat(selectedCert.map(data => (data["Temperatura do Local da Calibração"])).toString()));
                    workbook.sheet(0).cell('AV51').value(parseFloat(selectedCert.map(data => (data["Umidade Relativa do Local da Calibração"])).toString()));

                    workbook.sheet(0).cell('AT54').value((selectedCert.map(data => data["Padrão 1"])).toString());
                    workbook.sheet(0).cell('AT55').value((selectedCert.map(data => data["Padrão 2"])).toString());
                    workbook.sheet(0).cell('AT56').value((selectedCert.map(data => data["Padrão 3"])).toString());

                    workbook.sheet(0).cell('AV60').value((selectedCert.map(data => data["Data de Calibração"])).toString());
                    workbook.sheet(0).cell('AV61').value((selectedCert.map(data => data["Próxima Data de Calibração"])).toString());
                    workbook.sheet(0).cell('AV62').value("François Fernandez");
                    break;

                  case "R05150707I-3":
                    workbook.sheet(0).cell('AW3').value((selectedCert.map(data => data.Ano)).toString());
                    workbook.sheet(0).cell('AX3').value((selectedCert.map(data => data.Mês)).toString());
                    workbook.sheet(0).cell('AY3').value((selectedCert.map(data => data.PrecisoID)).toString());
                    workbook.sheet(0).cell('AW4').value((selectedCert.map(data => data.Incremental)).toString());

                    workbook.sheet(0).cell('AW5').value((selectedCert.map(data => data.Empresa)).toString());
                    workbook.sheet(0).cell('AW6').value((selectedCert.map(data => data.Endereço)).toString());
                    workbook.sheet(0).cell('AW7').value((selectedCert.map(data => data["Cidade c/ Estado"])).toString());
                    workbook.sheet(0).cell('AW8').value((selectedCert.map(data => data.Instrumento)).toString());
                    workbook.sheet(0).cell('AW9').value((selectedCert.map(data => data.Equipamento)).toString());

                    workbook.sheet(0).cell('AW10').value((selectedCert.map(data => data["Marca do Instrumento"])).toString());
                    workbook.sheet(0).cell('AW11').value((selectedCert.map(data => data["Modelo do Instrumento"])).toString());
                    workbook.sheet(0).cell('AW12').value((selectedCert.map(data => data["Classe do Instrumento"])).toString());
                    workbook.sheet(0).cell('AW13').value((selectedCert.map(data => data["Número de Série do Instrumento"])).toString());
                    workbook.sheet(0).cell('AW14').value((selectedCert.map(data => data["Identificação do Instrumento"])).toString());

                    workbook.sheet(0).cell('AW15').value(parseFloat(selectedCert.map(data => (data["Temperatura de Entrada - Inicio de Escala"])).toString()));
                    workbook.sheet(0).cell('AY15').value(parseFloat(selectedCert.map(data => (data["Temperatura de Entrada - Final de Escala"])).toString()));
                    workbook.sheet(0).cell('AW16').value(parseFloat(selectedCert.map(data => (data["Temperatura de Entrada - Inicio da Faixa de Uso"])).toString()));
                    workbook.sheet(0).cell('AY16').value(parseFloat(selectedCert.map(data => (data["Temperatura de Entrada - Final da Faixa de Uso"])).toString()));
                    workbook.sheet(0).cell('AW17').value(parseFloat(selectedCert.map(data => (data["Temperatura de Entrada - Valor de uma Divisão"])).toString()));

                    workbook.sheet(0).cell('AW18').value(parseFloat(selectedCert.map(data => (data["Temperatura de Saida - Inicio de Escala"])).toString()));
                    workbook.sheet(0).cell('AY18').value(parseFloat(selectedCert.map(data => (data["Temperatura de Saida - Final de Escala"])).toString()));
                    workbook.sheet(0).cell('AW19').value(parseFloat(selectedCert.map(data => (data["Temperatura de Saida - Inicio da Faixa de Uso"])).toString()));
                    workbook.sheet(0).cell('AY19').value(parseFloat(selectedCert.map(data => (data["Temperatura de Saida - Final da Faixa de Uso"])).toString()));
                    workbook.sheet(0).cell('AW20').value(parseFloat(selectedCert.map(data => (data["Temperatura de Saida - Valor de uma Divisão"])).toString()));

                    workbook.sheet(0).cell('AW21').value(parseFloat(selectedCert.map(data => (data["Umidade Relativa - Inicio de Escala"])).toString()));
                    workbook.sheet(0).cell('AY21').value(parseFloat(selectedCert.map(data => (data["Umidade Relativa - Final de Escala"])).toString()));
                    workbook.sheet(0).cell('AW22').value(parseFloat(selectedCert.map(data => (data["Umidade Relativa - Inicio da Faixa de Uso"])).toString()));
                    workbook.sheet(0).cell('AY22').value(parseFloat(selectedCert.map(data => (data["Umidade Relativa - Final da Faixa de Uso"])).toString()));
                    workbook.sheet(0).cell('AW23').value(parseFloat(selectedCert.map(data => (data["Umidade Relativa - Valor de uma Divisão"])).toString()));

                    workbook.sheet(0).cell('AU33').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 1"])).toString()));
                    workbook.sheet(0).cell('AU34').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 1"])).toString()));
                    workbook.sheet(0).cell('AU35').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 1"])).toString()));
                    workbook.sheet(0).cell('AU36').value(parseFloat(selectedCert.map(data => (data["V.V.C 4 - Leitura 1"])).toString()));
                    workbook.sheet(0).cell('AU37').value(parseFloat(selectedCert.map(data => (data["V.V.C 5 - Leitura 1"])).toString()));
                    workbook.sheet(0).cell('AV33').value(parseFloat(selectedCert.map(data => (data["V.I.I 1 - Leitura 1 - Linha 1"])).toString()));
                    workbook.sheet(0).cell('AW33').value(parseFloat(selectedCert.map(data => (data["V.I.I 2 - Leitura 1 - Linha 1"])).toString()));
                    workbook.sheet(0).cell('AX33').value(parseFloat(selectedCert.map(data => (data["V.I.I 3 - Leitura 1 - Linha 1"])).toString()));
                    workbook.sheet(0).cell('AV34').value(parseFloat(selectedCert.map(data => (data["V.I.I 1 - Leitura 1 - Linha 2"])).toString()));
                    workbook.sheet(0).cell('AW34').value(parseFloat(selectedCert.map(data => (data["V.I.I 2 - Leitura 1 - Linha 2"])).toString()));
                    workbook.sheet(0).cell('AX34').value(parseFloat(selectedCert.map(data => (data["V.I.I 3 - Leitura 1 - Linha 2"])).toString()));
                    workbook.sheet(0).cell('AV35').value(parseFloat(selectedCert.map(data => (data["V.I.I 1 - Leitura 1 - Linha 3"])).toString()));
                    workbook.sheet(0).cell('AW35').value(parseFloat(selectedCert.map(data => (data["V.I.I 2 - Leitura 1 - Linha 3"])).toString()));
                    workbook.sheet(0).cell('AX35').value(parseFloat(selectedCert.map(data => (data["V.I.I 3 - Leitura 1 - Linha 3"])).toString()));
                    workbook.sheet(0).cell('AV36').value(parseFloat(selectedCert.map(data => (data["V.I.I 1 - Leitura 1 - Linha 4"])).toString()));
                    workbook.sheet(0).cell('AW36').value(parseFloat(selectedCert.map(data => (data["V.I.I 2 - Leitura 1 - Linha 4"])).toString()));
                    workbook.sheet(0).cell('AX36').value(parseFloat(selectedCert.map(data => (data["V.I.I 3 - Leitura 1 - Linha 4"])).toString()));
                    workbook.sheet(0).cell('AV37').value(parseFloat(selectedCert.map(data => (data["V.I.I 1 - Leitura 1 - Linha 5"])).toString()));
                    workbook.sheet(0).cell('AW37').value(parseFloat(selectedCert.map(data => (data["V.I.I 2 - Leitura 1 - Linha 5"])).toString()));
                    workbook.sheet(0).cell('AX37').value(parseFloat(selectedCert.map(data => (data["V.I.I 3 - Leitura 1 - Linha 5"])).toString()));

                    workbook.sheet(0).cell('AU43').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 2"])).toString()));
                    workbook.sheet(0).cell('AU44').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 2"])).toString()));
                    workbook.sheet(0).cell('AU45').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 2"])).toString()));
                    workbook.sheet(0).cell('AU46').value(parseFloat(selectedCert.map(data => (data["V.V.C 4 - Leitura 2"])).toString()));
                    workbook.sheet(0).cell('AU47').value(parseFloat(selectedCert.map(data => (data["V.V.C 5 - Leitura 2"])).toString()));
                    workbook.sheet(0).cell('AV43').value(parseFloat(selectedCert.map(data => (data["V.I.I 1 - Leitura 2 - Linha 1"])).toString()));
                    workbook.sheet(0).cell('AW43').value(parseFloat(selectedCert.map(data => (data["V.I.I 2 - Leitura 2 - Linha 1"])).toString()));
                    workbook.sheet(0).cell('AX43').value(parseFloat(selectedCert.map(data => (data["V.I.I 3 - Leitura 2 - Linha 1"])).toString()));
                    workbook.sheet(0).cell('AV44').value(parseFloat(selectedCert.map(data => (data["V.I.I 1 - Leitura 2 - Linha 2"])).toString()));
                    workbook.sheet(0).cell('AW44').value(parseFloat(selectedCert.map(data => (data["V.I.I 2 - Leitura 2 - Linha 2"])).toString()));
                    workbook.sheet(0).cell('AX44').value(parseFloat(selectedCert.map(data => (data["V.I.I 3 - Leitura 2 - Linha 2"])).toString()));
                    workbook.sheet(0).cell('AV45').value(parseFloat(selectedCert.map(data => (data["V.I.I 1 - Leitura 2 - Linha 3"])).toString()));
                    workbook.sheet(0).cell('AW45').value(parseFloat(selectedCert.map(data => (data["V.I.I 2 - Leitura 2 - Linha 3"])).toString()));
                    workbook.sheet(0).cell('AX45').value(parseFloat(selectedCert.map(data => (data["V.I.I 3 - Leitura 2 - Linha 3"])).toString()));
                    workbook.sheet(0).cell('AV46').value(parseFloat(selectedCert.map(data => (data["V.I.I 1 - Leitura 2 - Linha 4"])).toString()));
                    workbook.sheet(0).cell('AW46').value(parseFloat(selectedCert.map(data => (data["V.I.I 2 - Leitura 2 - Linha 4"])).toString()));
                    workbook.sheet(0).cell('AX46').value(parseFloat(selectedCert.map(data => (data["V.I.I 3 - Leitura 2 - Linha 4"])).toString()));
                    workbook.sheet(0).cell('AV47').value(parseFloat(selectedCert.map(data => (data["V.I.I 1 - Leitura 2 - Linha 5"])).toString()));
                    workbook.sheet(0).cell('AW47').value(parseFloat(selectedCert.map(data => (data["V.I.I 2 - Leitura 2 - Linha 5"])).toString()));
                    workbook.sheet(0).cell('AX47').value(parseFloat(selectedCert.map(data => (data["V.I.I 3 - Leitura 2 - Linha 5"])).toString()));

                    workbook.sheet(0).cell('AU53').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 3"])).toString()));
                    workbook.sheet(0).cell('AU54').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 3"])).toString()));
                    workbook.sheet(0).cell('AU55').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 3"])).toString()));
                    workbook.sheet(0).cell('AU56').value(parseFloat(selectedCert.map(data => (data["V.V.C 4 - Leitura 3"])).toString()));
                    workbook.sheet(0).cell('AU57').value(parseFloat(selectedCert.map(data => (data["V.V.C 5 - Leitura 3"])).toString()));
                    workbook.sheet(0).cell('AV53').value(parseFloat(selectedCert.map(data => (data["V.I.I 1 - Leitura 3 - Linha 1"])).toString()));
                    workbook.sheet(0).cell('AW53').value(parseFloat(selectedCert.map(data => (data["V.I.I 2 - Leitura 3 - Linha 1"])).toString()));
                    workbook.sheet(0).cell('AX53').value(parseFloat(selectedCert.map(data => (data["V.I.I 3 - Leitura 3 - Linha 1"])).toString()));
                    workbook.sheet(0).cell('AV54').value(parseFloat(selectedCert.map(data => (data["V.I.I 1 - Leitura 3 - Linha 2"])).toString()));
                    workbook.sheet(0).cell('AW54').value(parseFloat(selectedCert.map(data => (data["V.I.I 2 - Leitura 3 - Linha 2"])).toString()));
                    workbook.sheet(0).cell('AX54').value(parseFloat(selectedCert.map(data => (data["V.I.I 3 - Leitura 3 - Linha 2"])).toString()));
                    workbook.sheet(0).cell('AV55').value(parseFloat(selectedCert.map(data => (data["V.I.I 1 - Leitura 3 - Linha 3"])).toString()));
                    workbook.sheet(0).cell('AW55').value(parseFloat(selectedCert.map(data => (data["V.I.I 2 - Leitura 3 - Linha 3"])).toString()));
                    workbook.sheet(0).cell('AX55').value(parseFloat(selectedCert.map(data => (data["V.I.I 3 - Leitura 3 - Linha 3"])).toString()));
                    workbook.sheet(0).cell('AV56').value(parseFloat(selectedCert.map(data => (data["V.I.I 1 - Leitura 3 - Linha 4"])).toString()));
                    workbook.sheet(0).cell('AW56').value(parseFloat(selectedCert.map(data => (data["V.I.I 2 - Leitura 3 - Linha 4"])).toString()));
                    workbook.sheet(0).cell('AX56').value(parseFloat(selectedCert.map(data => (data["V.I.I 3 - Leitura 3 - Linha 4"])).toString()));
                    workbook.sheet(0).cell('AV57').value(parseFloat(selectedCert.map(data => (data["V.I.I 1 - Leitura 3 - Linha 5"])).toString()));
                    workbook.sheet(0).cell('AW57').value(parseFloat(selectedCert.map(data => (data["V.I.I 2 - Leitura 3 - Linha 5"])).toString()));
                    workbook.sheet(0).cell('AX57').value(parseFloat(selectedCert.map(data => (data["V.I.I 3 - Leitura 3 - Linha 5"])).toString()));

                    workbook.sheet(0).cell('AW60').value((selectedCert.map(data => data["Local da Calibração"])).toString());
                    workbook.sheet(0).cell('AW61').value(parseFloat(selectedCert.map(data => (data["Temperatura do Local da Calibração"])).toString()));
                    workbook.sheet(0).cell('AW62').value(parseFloat(selectedCert.map(data => (data["Umidade Relativa do Local da Calibração"])).toString()));

                    workbook.sheet(0).cell('AU65').value((selectedCert.map(data => data["Padrão 1"])).toString());
                    workbook.sheet(0).cell('AU66').value((selectedCert.map(data => data["Padrão 2"])).toString());
                    workbook.sheet(0).cell('AU67').value((selectedCert.map(data => data["Padrão 3"])).toString());

                    workbook.sheet(0).cell('AW72').value((selectedCert.map(data => data["Data de Calibração"])).toString());
                    workbook.sheet(0).cell('AW73').value((selectedCert.map(data => data["Próxima Data de Calibração"])).toString());
                    workbook.sheet(0).cell('AW74').value("François Fernandez");
                    break;

                  case "R100180915F":
                    workbook.sheet(0).cell('AV3').value((selectedCert.map(data => data.Ano)).toString());
                    workbook.sheet(0).cell('AW3').value((selectedCert.map(data => data.Mês)).toString());
                    workbook.sheet(0).cell('AX3').value((selectedCert.map(data => data.PrecisoID)).toString());
                    workbook.sheet(0).cell('AV4').value((selectedCert.map(data => data.Incremental)).toString());

                    workbook.sheet(0).cell('AV5').value((selectedCert.map(data => data.Empresa)).toString());
                    workbook.sheet(0).cell('AV6').value((selectedCert.map(data => data.Endereço)).toString());
                    workbook.sheet(0).cell('AV7').value((selectedCert.map(data => data["Cidade c/ Estado"])).toString());

                    workbook.sheet(0).cell('AV8').value((selectedCert.map(data => data.Instrumento)).toString());
                    workbook.sheet(0).cell('AV9').value((selectedCert.map(data => data.Equipamento)).toString());
                    workbook.sheet(0).cell('AV10').value((selectedCert.map(data => data["Marca do Instrumento"])).toString());
                    workbook.sheet(0).cell('AV11').value((selectedCert.map(data => data["Modelo do Instrumento"])).toString());
                    workbook.sheet(0).cell('AV12').value((selectedCert.map(data => data["Número de Série do Instrumento"])).toString());
                    workbook.sheet(0).cell('AV13').value((selectedCert.map(data => data["Identificação do Instrumento"])).toString());

                    workbook.sheet(0).cell('AV14').value(parseFloat(selectedCert.map(data => (data["Volume Nominal"])).toString()));
                    workbook.sheet(0).cell('AV15').value(parseFloat(selectedCert.map(data => (data["Inicio da Faixa de Uso"])).toString()));
                    workbook.sheet(0).cell('AX15').value(parseFloat(selectedCert.map(data => (data["Final da Faixa de Uso"])).toString()));
                    workbook.sheet(0).cell('AX14').value((selectedCert.map(data => data.Unidade)).toString());
                    workbook.sheet(0).cell('AV16').value(parseFloat(selectedCert.map(data => (data["Valor de uma Divisao"])).toString()));

                    workbook.sheet(0).cell('AT21').value(parseFloat(selectedCert.map(data => (data["V.I.I 1 - Leitura 1"])).toString()));
                    workbook.sheet(0).cell('AT22').value(parseFloat(selectedCert.map(data => (data["V.I.I 2 - Leitura 1"])).toString()));
                    workbook.sheet(0).cell('AT23').value(parseFloat(selectedCert.map(data => (data["V.I.I 3 - Leitura 1"])).toString()));
                    workbook.sheet(0).cell('AT24').value(parseFloat(selectedCert.map(data => (data["V.I.I 4 - Leitura 1"])).toString()));
                    workbook.sheet(0).cell('AT25').value(parseFloat(selectedCert.map(data => (data["V.I.I 5 - Leitura 1"])).toString()));

                    workbook.sheet(0).cell('AU21').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 1 - Linha 1"])).toString()));
                    workbook.sheet(0).cell('AV21').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 1 - Linha 1"])).toString()));
                    workbook.sheet(0).cell('AW21').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 1 - Linha 1"])).toString()));
                    workbook.sheet(0).cell('AU22').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 1 - Linha 2"])).toString()));
                    workbook.sheet(0).cell('AV22').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 1 - Linha 2"])).toString()));
                    workbook.sheet(0).cell('AW22').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 1 - Linha 2"])).toString()));
                    workbook.sheet(0).cell('AU23').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 1 - Linha 3"])).toString()));
                    workbook.sheet(0).cell('AV23').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 1 - Linha 3"])).toString()));
                    workbook.sheet(0).cell('AW23').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 1 - Linha 3"])).toString()));
                    workbook.sheet(0).cell('AU24').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 1 - Linha 4"])).toString()));
                    workbook.sheet(0).cell('AV24').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 1 - Linha 4"])).toString()));
                    workbook.sheet(0).cell('AW24').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 1 - Linha 4"])).toString()));
                    workbook.sheet(0).cell('AU25').value(parseFloat(selectedCert.map(data => (data["V.V.C 1 - Leitura 1 - Linha 5"])).toString()));
                    workbook.sheet(0).cell('AV25').value(parseFloat(selectedCert.map(data => (data["V.V.C 2 - Leitura 1 - Linha 5"])).toString()));
                    workbook.sheet(0).cell('AW25').value(parseFloat(selectedCert.map(data => (data["V.V.C 3 - Leitura 1 - Linha 5"])).toString()));

                    workbook.sheet(0).cell('AV28').value((selectedCert.map(data => (data["Local da Calibração"])).toString()));
                    workbook.sheet(0).cell('AV29').value(parseFloat(selectedCert.map(data => (data["Temperatura do Local da Calibração"])).toString()));
                    workbook.sheet(0).cell('AV30').value(parseFloat(selectedCert.map(data => (data["Umidade Relativa do Local da Calibração"])).toString()));
                    workbook.sheet(0).cell('AV31').value(parseFloat(selectedCert.map(data => (data["Pressão Atmosférica"])).toString()));

                    workbook.sheet(0).cell('AT33').value((selectedCert.map(data => data["Padrão 1"])).toString());
                    workbook.sheet(0).cell('AT34').value((selectedCert.map(data => data["Padrão 2"])).toString());
                    workbook.sheet(0).cell('AT35').value((selectedCert.map(data => data["Padrão 3"])).toString());

                    workbook.sheet(0).cell('AV37').value((selectedCert.map(data => data["Data de Calibração"])).toString());
                    workbook.sheet(0).cell('AV38').value((selectedCert.map(data => data["Próxima Data de Calibração"])).toString());
                    workbook.sheet(0).cell('AV39').value("François Fernandez");
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
              }
            )}
          )}
      }   
    }
    xmlhttp.send(); 
  }
}
