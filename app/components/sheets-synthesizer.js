/*globals XLSX*/

import Ember from "ember";
import EmberUploader from "ember-uploader";

export default EmberUploader.FileField.extend({
  multiple: true,

  cellMaps: {
    "A.P.EX.02.01 - Rev 13.1 - RODOVIARIO": {
      revVersion: 13.1,
      driverName: "J6",
      startDate: "J9",
      finalDate: "P9",
      hoursAdc: "AP8",
      hoursEsp: "AP10",
      hoursExt: "AP9",
      daysFinished: "AS8",
      daysPendent: "AS9"
    },
    "A.P.EX.02.01 - Rev 13- RODOVIARIO": {
      revVersion: 13,
      driverName: "J6",
      startDate: "J9",
      finalDate: "O9",
      hoursAdc: "AP8",
      hoursEsp: "AP10",
      hoursExt: "AP9",
      daysFinished: "AS8",
      daysPendent: "AS9"
    },
    "A.P.EX.02.01 - Rev 12 - RODOVIARIO": {
      revVersion: 12,
      driverName: "J6",
      startDate: "J9",
      finalDate: "O9",
      hoursAdc: "AF8",
      hoursEsp: "AF10",
      hoursExt: "AF9",
      daysFinished: "AI8",
      daysPendent: "AI9"
    },
    "A.P.EX.02.01 - Rev 11.01 - RODOVIARIO": {
      revVersion: 11.01,
      driverName: "J6",
      startDate: "J9",
      finalDate: "O9",
      hoursAdc: "AF8",
      hoursEsp: "AF10",
      hoursExt: "AF9",
      daysFinished: "AI8",
      daysPendent: "AI9"
    },
    "A.P.EX.02.01 - Rev 11 - RODOVIARIO": {
      revVersion: 11,
      driverName: "J6",
      startDate: "J9",
      finalDate: "O9",
      hoursAdc: "AC8",
      hoursEsp: "AC10",
      hoursExt: "AC9",
      daysFinished: "AF8",
      daysPendent: "AF9"
    },
    "A.P.EX.02.01 - Rev 11 - CARVAO": {
      revVersion: 11,
      driverName: "I6",
      startDate: "I9",
      finalDate: "O9",
      hoursAdc: "AD8",
      hoursEsp: "AD10",
      hoursExt: "AD9",
      daysFinished: "AG8",
      daysPendent: "AG9"
    },
    "A.P.EX.02.01 - Rev 10.3 - RODOVIARIO": {
      revVersion: 10.3,
      driverName: "H6",
      startDate: "H9",
      finalDate: "M9",
      hoursAdc: "AA8",
      hoursEsp: "AA10",
      hoursExt: "AA9"
    },
    "A.P.EX.02.01 - Rev 10.3 - CARVAO": {
      revVersion: 10.3,
      driverName: "G6",
      startDate: "G9",
      finalDate: "M9",
      hoursAdc: "AB8",
      hoursEsp: "AB10",
      hoursExt: "AB9"
    },
    "Rev 9": {
      revVersion: 9,
      driverName: "G6",
      startDate: "G9",
      finalDate: "M9",
      hoursAdc: "X8",
      hoursEsp: "X10",
      hoursExt: "X9"
    },
    "A.P.EX.02.01 - Rev S/V - RODOVIARIO": {
      revVersion: 8,
      driverName: "H6",
      startDate: "H9",
      finalDate: "L9",
      hoursAdc: "V8",
      hoursEsp: "V10",
      hoursExt: "V9"
    },
    "A.P.EX.02.01 - Rev S/V - CARVAO": {
      revVersion: 8,
      driverName: "G6",
      startDate: "G9",
      finalDate: "L9",
      hoursAdc: "V8",
      hoursEsp: "V10",
      hoursExt: "V8"
    }
  },

  filesDidChange: function(files) {
    var j,
      len,
      workbook,
      self = this;

    var translateSheet = function(e) {
      workbook = XLSX.read(e.target.result, { type: "binary" });

      //TODO: REMOVE
      console.log("Sheet " + e.target.fileName);
      console.log(workbook);

      var version = workbook.Props.Title;
      if (version === undefined) {
        if (workbook.Sheets["JORNADA DE MOTORISTA"]["G6"] === undefined) {
          version = "A.P.EX.02.01 - Rev S/V - RODOVIARIO";
        } else {
          version = "A.P.EX.02.01 - Rev S/V - CARVAO";
        }
      }

      if (
        workbook.Sheets["JORNADA DE MOTORISTA"]["AJ1"] &&
        workbook.Sheets["JORNADA DE MOTORISTA"]["AJ1"].v ===
          "RODOVIARIO - REV 11"
      ) {
        version = "A.P.EX.02.01 - Rev 11.01 - RODOVIARIO";
      }

      console.log("Version:", version);

      var cellMap = self.get("cellMaps")[version];
      var data;

      console.log("Cell Map:", cellMap);

      if (cellMap === undefined) {
        //TODO
        alert(
          'Versão da Planilha "' +
            version +
            '" não Cadastrada: ' +
            e.target.fileName
        );
      } else if (
        (workbook.Sheets["JORNADA DE MOTORISTA"][cellMap.hoursAdc] || {}).w ===
          "#VALUE!" ||
        (workbook.Sheets["JORNADA DE MOTORISTA"][cellMap.hoursEsp] || {}).w ===
          "#VALUE!" ||
        (workbook.Sheets["JORNADA DE MOTORISTA"][cellMap.hoursExt] || {}).w ===
          "#VALUE!"
      ) {
        alert("Erro de Fórmula - Planilha não importada: " + e.target.fileName);
      } else if (
        workbook.Sheets["JORNADA DE MOTORISTA"][cellMap.driverName] ===
        undefined
      ) {
        alert(
          "Sem Nome do Motorista - Planilha não importada: " + e.target.fileName
        );
      } else if (cellMap.revVersion >= 11) {
        data = {
          fileName: e.target.fileName,
          version: version,
          driverName:
            workbook.Sheets["JORNADA DE MOTORISTA"][cellMap.driverName].v,
          startDate:
            workbook.Sheets["JORNADA DE MOTORISTA"][cellMap.startDate].w,
          finalDate:
            workbook.Sheets["JORNADA DE MOTORISTA"][cellMap.finalDate].w,
          hoursAdc: workbook.Sheets["JORNADA DE MOTORISTA"][cellMap.hoursAdc].v,
          hoursEsp: workbook.Sheets["JORNADA DE MOTORISTA"][cellMap.hoursEsp].v,
          hoursExt: workbook.Sheets["JORNADA DE MOTORISTA"][cellMap.hoursExt].v,
          daysFinished:
            workbook.Sheets["JORNADA DE MOTORISTA"][cellMap.daysFinished].v,
          daysPendent:
            workbook.Sheets["JORNADA DE MOTORISTA"][cellMap.daysPendent].v
        };

        self.sheetsData.pushObject(data);
      } else {
        data = {
          fileName: e.target.fileName,
          version: version,
          driverName:
            workbook.Sheets["JORNADA DE MOTORISTA"][cellMap.driverName].v,
          startDate:
            workbook.Sheets["JORNADA DE MOTORISTA"][cellMap.startDate].w,
          finalDate:
            workbook.Sheets["JORNADA DE MOTORISTA"][cellMap.finalDate].w,
          hoursAdc: workbook.Sheets["JORNADA DE MOTORISTA"][cellMap.hoursAdc].v,
          hoursEsp: workbook.Sheets["JORNADA DE MOTORISTA"][cellMap.hoursEsp].v,
          hoursExt: workbook.Sheets["JORNADA DE MOTORISTA"][cellMap.hoursExt].v
        };

        self.sheetsData.pushObject(data);

        //TODO: REMOVE
        console.log(data);
      }

      //TODO: REMOVE
      //console.log('-------------------- END');
    };

    //TODO: REMOVE
    //console.log(this.get('cellMaps'));

    this.set("sheetsData", Ember.A());

    for (j = 0, len = files.length; j < len; j++) {
      var file = files[j];

      var reader = new FileReader();
      reader.fileName = file.name;
      reader.onload = translateSheet;
      reader.readAsBinaryString(file);
    }
  }
});
