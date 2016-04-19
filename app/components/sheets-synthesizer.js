/*globals XLSX*/
import Ember from 'ember';
import EmberUploader from 'ember-uploader';

export default EmberUploader.FileField.extend({
  multiple: true,

  cellMaps: {
    'A.P.EX.02.01 - Rev 10.3 - RODOVIARIO' : {
      driverName: 'H6',
      startDate:  'H9',
      finalDate:  'M9',
      hoursAdc:   'AA8',
      hoursEsp:   'AA10',
      hoursExt:   'AA9'
    },
    'A.P.EX.02.01 - Rev 10.3 - CARVAO' : {
      driverName: 'G6',
      startDate:  'G9',
      finalDate:  'M9',
      hoursAdc:   'AB8',
      hoursEsp:   'AB10',
      hoursExt:   'AB9'
    },
    'Rev 9' : {
      driverName: 'G6',
      startDate:  'G9',
      finalDate:  'M9',
      hoursAdc:   'X8',
      hoursEsp:   'X10',
      hoursExt:   'X9'
    },
    'A.P.EX.02.01 - Rev 8 - RODOVIARIO' : {
      driverName: 'H6',
      startDate:  'H9',
      finalDate:  'L9',
      hoursAdc:   'V8',
      hoursEsp:   'V10',
      hoursExt:   'V9'
    },
    'A.P.EX.02.01 - Rev 8 - CARVAO' : {
      driverName: 'G6',
      startDate:  'G9',
      finalDate:  'L9',
      hoursAdc:   'V8',
      hoursEsp:   'V10',
      hoursExt:   'V8'
    }
  },

  filesDidChange: function(files) {
    var j, len, workbook, self = this;

    //TODO: REMOVE
    //console.log(this.get('cellMaps'));

    this.set('sheetsData', Ember.A());

    for (j = 0, len = files.length; j < len; j++) {
      var file = files[j];

      var reader = new FileReader();
      reader.fileName = file.name;
      reader.onload = function(e) {
        workbook = XLSX.read(e.target.result, {type: 'binary'});

        //TODO: REMOVE
        //console.log('Sheet '+e.target.fileName);
        //console.log(workbook);

        var version = workbook.Props.Title;
        if (version === undefined) {
          if (workbook.Sheets['JORNADA DE MOTORISTA']['G6'] === undefined) {
            version = 'A.P.EX.02.01 - Rev 8 - RODOVIARIO';
          } else {
            version = 'A.P.EX.02.01 - Rev 8 - CARVAO';
          }
        }

        var cellMap = self.get('cellMaps')[version];

        if (cellMap === undefined) {
          //TODO
          alert('Versão da Planilha "'+version+'" não Cadastrada: '+e.target.fileName);
        } else if (workbook.Sheets['JORNADA DE MOTORISTA'][cellMap.hoursAdc].w == "#VALUE!" || workbook.Sheets['JORNADA DE MOTORISTA'][cellMap.hoursEsp].w == "#VALUE!" ||  workbook.Sheets['JORNADA DE MOTORISTA'][cellMap.hoursExt].w == "#VALUE!") {
          alert('Planilha com Erro de Fórmula não importada: '+e.target.fileName);
        } else {
          var data = {
            fileName: e.target.fileName,
            version:    version,
            driverName: workbook.Sheets['JORNADA DE MOTORISTA'][cellMap.driverName].v,
            startDate:  workbook.Sheets['JORNADA DE MOTORISTA'][cellMap.startDate].w,
            finalDate:  workbook.Sheets['JORNADA DE MOTORISTA'][cellMap.finalDate].w,
            hoursAdc:   workbook.Sheets['JORNADA DE MOTORISTA'][cellMap.hoursAdc].v,
            hoursEsp:   workbook.Sheets['JORNADA DE MOTORISTA'][cellMap.hoursEsp].v,
            hoursExt:   workbook.Sheets['JORNADA DE MOTORISTA'][cellMap.hoursExt].v
          };

          self.sheetsData.pushObject(data);

          //TODO: REMOVE
          //console.log(data);
        }

        //TODO: REMOVE
        //console.log('-------------------- END');
      };
      reader.readAsBinaryString(file);

    }
  }
});
