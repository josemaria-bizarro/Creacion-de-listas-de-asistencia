const SheetDATA = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DATA");
const SheetCARATULA = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CARATULA"); 
var Indexlsr =  SheetCARATULA.getRange("A15:A").getDisplayValues();

const IDPlantilla = SheetDB.getRange("X2").getDisplayValue();
const IDFolderDestino = SheetDB.getRange("BW2").getDisplayValue();
var TotalList = Indexlsr.filter(String).length;


function GetDataList() 
{
const data = SheetDATA.getRange(2,1,121,91).getDisplayValues();
data.forEach(row => { 
          CreateLista(IDPlantilla,IDFolderDestino,
                    row[0],row[1],row[2],row[3],row[4],row[5],row[6],row[7],row[8],row[9],
                    row[10],row[11],row[12],row[13],row[14],row[15],row[16],row[17],row[18],row[19],
                    row[20],row[21],row[22],row[23],row[24],row[25],row[26],row[27],row[28],row[29],
                    row[30],row[31],row[32],row[33],row[34],row[35],row[36],row[37],row[38],row[39],
                    row[40],row[41],row[42],row[43],row[44],row[45],row[46],row[47],row[48],row[49],
                    row[50],row[51],row[52],row[53],row[54],row[55],row[56],row[57],row[58],row[59],
                    row[60],row[61],row[62],row[63],row[64],row[65],row[66],row[67],row[68],row[69],
                    row[70],row[71],row[72],row[73],row[74],row[75],row[76],row[77],row[78],row[79],
                    row[80],row[81],row[82],row[83],row[84],row[85],row[86],row[87],row[88],row[89],row[90],row[91],row[92]
                    );     
  });
}

function CreateLista(IDPlantilla,IDFolderDestino,
  INLIS,CATEG,OPEDU,IDOPE,PLANT,CIESC,PERIO,INPER,FNPER,PROFE,
  EMPROF,IDPROF,ASINT,ASLEG,IDASI,PONEV,PONEX,SEMES,TURNO,GRUPO,
  LISNM,IDSEM,IDTURN,IDAÑO,DATA1,DATA2,DAULA1,DAULA2,DAULA3,DAULA4,
  DAULA5,AL1,AL2,AL3,AL4,AL5,AL6,AL7,AL8,AL9,AL10,
  AL11,AL12,AL13,AL14,AL15,AL16,AL17,AL18,AL19,AL20,
  AL21,AL22,AL23,AL24,AL25,AL26,AL27,AL28,AL29,AL30,
  AL31,AL32,AL33,AL34,AL35,AL36,AL37,AL38,AL39,AL40,
  AL41,AL42,AL43,AL44,AL45,AL46,AL47,AL48,AL49,AL50,
  AL51,AL52,AL53,AL54,AL55,AL56,AL57,AL58,AL59,AL60
  )
  {

 const SheetFile = DriveApp.getFileById(IDPlantilla);
 const FolderConte = DriveApp.getFolderById(IDFolderDestino);

 
 if (INLIS <= TotalList){
   
   switch (CATEG)
   {
     case "DATA":
          var CurDate = Utilities.formatDate(new Date(), "GMT-6", "dd/MM/yyyy");
          var TempAlum = [[AL1],[AL2],[AL3],[AL4],[AL5],[AL6],[AL7],[AL8],[AL9],[AL10],
                          [AL11],[AL12],[AL13],[AL14],[AL15],[AL16],[AL17],[AL18],[AL19],[AL20],
                          [AL21],[AL22],[AL23],[AL24],[AL25],[AL26],[AL27],[AL28],[AL29],[AL30],
                          [AL31],[AL32],[AL33],[AL34],[AL35],[AL36],[AL37],[AL38],[AL39],[AL40],
                          [AL41],[AL42],[AL43],[AL44],[AL45],[AL46],[AL47],[AL48],[AL49],[AL50],
                          [AL51],[AL52],[AL53],[AL54],[AL55],[AL56],[AL57],[AL58],[AL59],[AL60]]
              
          Alumnos = TempAlum.filter(String)
          
          TotalAlum = TempAlum.filter(String).length;
          
          //26-05-23 firma profesor
          if (TotalAlum>30)
          {
            var TotalAlum1 = 30
            var TotalAlum2 = Math.round(TotalAlum-30)
          }
          else
          {
            var TotalAlum1 = TotalAlum
            var TotalAlum2 = 0
          }  
        
          
          
          const ListCopia = SheetFile.makeCopy(FolderConte);
          var Idlis = ListCopia.getId();
          var Cache = CacheService.getDocumentCache().put("ID",Idlis);

          var Lista = SpreadsheetApp.openById(Idlis);
          var Hoja = Lista.getSheetByName("ORIGEN");
          

              Hoja.getRange("B6").setValue(ASLEG);
              Hoja.getRange("B8").setValue(ASINT);
              Hoja.getRange("D8").setValue(IDASI);
              Hoja.getRange("F2").setValue(PROFE);
              Hoja.getRange("AH1").setValue(PERIO);
              Hoja.getRange("BA1").setValue(PLANT);
              Hoja.getRange("BU1").setValue(TURNO);
              Hoja.getRange("AH3").setValue(INPER);
              Hoja.getRange("AH4").setValue(FNPER);
              Hoja.getRange("BA3").setValue(CIESC);
              Hoja.getRange("AV4").setValue(GRUPO);
              Hoja.getRange("BL3").setValue(SEMES);
              Hoja.getRange("BY5").setValue(PONEV);
              Hoja.getRange("CA5").setValue(PONEX);

              Hoja.getRange("CA1").setValue("Creación " + CurDate);

              //Hoja.getRange(10,3,TotalAlum,1).setValues(Alumnos);//26-05-23 firma profesor
              var Alumnos1 =[]
              for (var i=0; i<TotalAlum1; i++)
              {
                Alumnos1.push(Alumnos[i])
              }
              try
              {
                  Hoja.getRange(10,3,TotalAlum1,1).setValues(Alumnos1);//26-05-23 firma profesor

                  //Hoja.getRange(10,3,TotalAlum,1).setValues(Alumnos);//26-05-23 firma profesor
                  
                  if (TotalAlum2 > 0)
                  {
                      let Alumnos2 =[]
                      for (var i=0; i<TotalAlum2; i++)
                      {
                        var k=i+30
                          Alumnos2.push(Alumnos[k])
                      }
                      Hoja.getRange(42,3,TotalAlum2,1).setValues(Alumnos2);//26-05-23 firma profesor
                  }
              } // fin TRY
              catch(err)
              {
                var errMsg=('NO HAY ALUMNOS SELECCIONADOS...')
                var html= HtmlService.createHtmlOutput(errMsg)
                  .setWidth(500)
                  .setHeight(100);
                SpreadsheetApp.getUi()
                  .showModalDialog(html,'ERROR');
              }
              var Evidencias = Hoja.getRange("BX10").setFormula('=IF(C10:C>0,IF(COUNT(F10:BW10)=0,"-",COUNT({G10,I10,K10,M10,O10,Q10,S10,U10,W10,Y10,AA10,AC10,AE10,AG10,AI10,AK10,AM10,AO10,AQ10,AS10,AU10,AW10,AY10,BA10,BC10,BE10,BG10,BI10,BK10,BM10,BO10,BQ10,BS10,BU10,BW10})),"")');
              //var FillEvidencias = Hoja.getRange(10,76,TotalAlum)//26-05-23 firma profesor
              var FillEvidencias = Hoja.getRange(10,76,TotalAlum1)//26-05-23 firma profesor

              
                    
              var EvaCont = Hoja.getRange("BY10").setFormula('=IF(COUNT(F10:BW10)=0,"",FIXED(SUM({G10,I10,K10,M10,O10,Q10,S10,U10,W10,Y10,AA10,AC10,AE10,AG10,AI10,AK10,AM10,AO10,AQ10,AS10,AU10,AW10,AY10,BA10,BC10,BE10,BG10,BI10,BK10,BM10,BO10,BQ10,BS10,BU10,BW10})/$BY$4,1))');
              //var FillEvaCont = Hoja.getRange(10,77,TotalAlum)//26-05-23 firma profesor
              var FillEvaCont = Hoja.getRange(10,77,TotalAlum1)//26-05-23 firma profesor
              

              var NoFaltas = Hoja.getRange("CF10").setFormula('=IF((COUNTA(F10:BW10))=0,"",FIXED(COUNTIF({F10,H10,J10,L10,N10,P10,R10,T10,V10,X10,Z10,AB10,AD10,AF10,AH10,AJ10,AL10,AN10,AP10,AR10,AT10,AV10,AX10,AZ10,BB10,BD10,BF10,BH10,BJ10,BL10,BN10,BP10,BR10,BT10,BV10},"=/"),))');
              //var FillNoFaltas = Hoja.getRange(10,84,TotalAlum)//26-05-23 firma profesor
              var FillNoFaltas = Hoja.getRange(10,84,TotalAlum1)//26-05-23 firma profesor
              
              //Hoja.getRange("BX10").copyTo(FillEvidencias);// 26-05-23 firma profesor
              //Hoja.getRange("BY10").copyTo(FillEvaCont);// 26-05-23 firma profesor
              //Hoja.getRange("CF10").copyTo(FillNoFaltas);// 26-05-23 firma profesor

              Hoja.getRange("BX10").copyTo(FillEvidencias);// 26-05-23 firma profesor
              Hoja.getRange("BY10").copyTo(FillEvaCont);// 26-05-23 firma profesor
              Hoja.getRange("CF10").copyTo(FillNoFaltas);// 26-05-23 firma profesor
              
              if (TotalAlum2>0)
                {
                  var FillEvidencias1 = Hoja.getRange(10,76,TotalAlum2)//26-05-23 firma profesor
                  var FillEvaCont1 = Hoja.getRange(10,77,TotalAlum2)//26-05-23 firma profesor
                  var FillNoFaltas1 = Hoja.getRange(10,84,TotalAlum2)//26-05-23 firma profesor
                  Hoja.getRange("BX42").copyTo(FillEvidencias1);// 26-05-23 firma profesor
                  Hoja.getRange("BY42").copyTo(FillEvaCont1);// 26-05-23 firma profesor
                  Hoja.getRange("CF42").copyTo(FillNoFaltas1);// 26-05-23 firma profesor
                }
              

              try
              {
              Lista.addEditor(EMPROF);
              }
              catch(err)
              {
                var errMsg=('Error al asignar profesor como editor %s',err.message)
                var html= HtmlService.createHtmlOutput(errMsg)
                  .setWidth(500)
                  .setHeight(100);
                SpreadsheetApp.getUi()
                  .showModalDialog(html,'FALLA');
              }

              var Bloqueo1 = Hoja.getRange('A1:E');
              var Bloqueo2 = Hoja.getRange('F1:CG5');
              var Bloqueo3 = Hoja.getRange('BX5:BZ');
              var Bloqueo4 = Hoja.getRange('CA5:CG9');
              var Bloqueo5 = Hoja.getRange('CB10:CC');
              var Bloqueo6 = Hoja.getRange('CE10:CG51');

              var Protection1 = Bloqueo1.protect();
              var Protection2 = Bloqueo2.protect();
              var Protection3 = Bloqueo3.protect();
              var Protection4 = Bloqueo4.protect();
              var Protection5 = Bloqueo5.protect();
              var Protection6 = Bloqueo6.protect();

              Protection1.removeEditor(EMPROF);
              Protection2.removeEditor(EMPROF);
              Protection3.removeEditor(EMPROF);
              Protection4.removeEditor(EMPROF);
              Protection5.removeEditor(EMPROF);
              Protection6.removeEditor(EMPROF);
              
            DriveApp.getFileById(Idlis).setName(LISNM);
            DriveApp.getFileById(Idlis).setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.EDIT);

     break;

     case "GRUPO":
            var Cache = CacheService.getDocumentCache().get("ID"); 
            var TempAlum = [[AL1],[AL2],[AL3],[AL4],[AL5],[AL6],[AL7],[AL8],[AL9],[AL10],
                          [AL11],[AL12],[AL13],[AL14],[AL15],[AL16],[AL17],[AL18],[AL19],[AL20],
                          [AL21],[AL22],[AL23],[AL24],[AL25],[AL26],[AL27],[AL28],[AL29],[AL30],
                          [AL31],[AL32],[AL33],[AL34],[AL35],[AL36],[AL37],[AL38],[AL39],[AL40],
                          [AL41],[AL42],[AL43],[AL44],[AL45],[AL46],[AL47],[AL48],[AL49],[AL50],
                          [AL51],[AL52],[AL53],[AL54],[AL55],[AL56],[AL57],[AL58],[AL59],[AL60]]
                          
            Alumnos = TempAlum.filter(String)
            TotalAlum = TempAlum.filter(String).length;

//26-05-23 firma profesor
          if (TotalAlum>30)
          {
            var TotalAlum1 = 30
            var TotalAlum2 = Math.round(TotalAlum-30)
          }
          else
          {
            var TotalAlum1 = TotalAlum
            var TotalAlum2 = 0
          }  


          var Alumnos1 =[]
          for (let i=0; i<TotalAlum1; i++)
          {
            Alumnos1.push(Alumnos[i])
          }
            var Lista = SpreadsheetApp.openById(Cache);
            var Hoja = Lista.getSheetByName("ORIGEN");
            Hoja.getRange(10,4,TotalAlum1,1).setValues(Alumnos1);//26-05-23 firma profesor

          //Hoja.getRange(10,4,TotalAlum,1).setValues(Alumnos);//26-05-23 firma profesor
          
          if (TotalAlum2 > 0)
          {
              let Alumnos2 =[]
              for (let i=0; i<TotalAlum2; i++)
              {
                let k=i+30
                  Alumnos2.push(Alumnos[k])
              }
              var Lista = SpreadsheetApp.openById(Cache);
              var Hoja = Lista.getSheetByName("ORIGEN");
              Hoja.getRange(42,4,TotalAlum2,1).setValues(Alumnos2);//26-05-23 firma profesor
          }


           // var Lista = SpreadsheetApp.openById(Cache);
           // var Hoja = Lista.getSheetByName("ORIGEN");
           //Hoja.getRange(10,4,TotalAlum,1).setValues(Alumnos);
     break;
     
     case "CORREO":
          var Cache = CacheService.getDocumentCache().get("ID");
          var TempAlum = [[AL1],[AL2],[AL3],[AL4],[AL5],[AL6],[AL7],[AL8],[AL9],[AL10],
                          [AL11],[AL12],[AL13],[AL14],[AL15],[AL16],[AL17],[AL18],[AL19],[AL20],
                          [AL21],[AL22],[AL23],[AL24],[AL25],[AL26],[AL27],[AL28],[AL29],[AL30],
                          [AL31],[AL32],[AL33],[AL34],[AL35],[AL36],[AL37],[AL38],[AL39],[AL40],
                          [AL41],[AL42],[AL43],[AL44],[AL45],[AL46],[AL47],[AL48],[AL49],[AL50],
                          [AL51],[AL52],[AL53],[AL54],[AL55],[AL56],[AL57],[AL58],[AL59],[AL60]]
                          
            Alumnos = TempAlum.filter(String)
            TotalAlum = TempAlum.filter(String).length;
//26-05-23 firma profesor
          if (TotalAlum>30)
          {
            var TotalAlum1 = 30
            var TotalAlum2 = Math.round(TotalAlum-30)
          }
          else
          {
            var TotalAlum1 = TotalAlum
            var TotalAlum2 = 0
          }  
             var Alumnos1 =[]
          for (let i=0; i<TotalAlum1; i++)
          {
            Alumnos1.push(Alumnos[i])
          }
          var Lista = SpreadsheetApp.openById(Cache);
          var Hoja = Lista.getSheetByName("ORIGEN");
          Hoja.getRange(10,5,TotalAlum1,1).setValues(Alumnos1);//26-05-23 firma profesor
          
          if (TotalAlum2 > 0)
          {
              let Alumnos2 =[]
              for (let i=0; i<TotalAlum2; i++)
              {
                let k=i+30
                  Alumnos2.push(Alumnos[k])
              }
              var Lista = SpreadsheetApp.openById(Cache);
              var Hoja = Lista.getSheetByName("ORIGEN");
              Hoja.getRange(42,5,TotalAlum2,1).setValues(Alumnos2);//26-05-23 firma profesor
          }

          // var Lista = SpreadsheetApp.openById(Cache);
          // var Hoja = Lista.getSheetByName("ORIGEN");

          //Hoja.getRange(10,5,TotalAlum,1).setValues(Alumnos);

     break;
   }
  }
  }

