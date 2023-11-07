const SheetDB = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB"); 
let opEduAnt =""

function bddata(){
 const catAcade = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1T5KQCvXiYsGewSoc_A8qugUwSs-ft1NNds4oaNgQxkw/edit#gid=378198808");
 const catAlumn = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1kzmgvrbkQ7ehUslRxR4ghkpSeI2yT99M3h-Dyo-EEVY/edit#gid=446758492");

var opEduActual=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CARATULA").getRange("F4").getDisplayValue()
 

    
    SheetDB.getRange("AC1:AK").clearContent(); 
    var asign = catAcade.getSheetByName("TABLA ASIGNATURAS");
    var d1asign = asign.getRange("A1:A").getValues(); //CLAVE
    var d2asign = asign.getRange("C1:C").getValues(); //NOMBRE INTERNO
    var d3asign = asign.getRange("I1:I").getValues(); //CUATRI
    var d4asign = asign.getRange("B1:B").getValues(); //NOMBRE LEGAL
    var d5asign = asign.getRange("J1:K").getValues(); //PONDERA
    var d6asign = asign.getRange("D1:D").getValues(); //OPCION EDU
    
       
      SheetDB.getRange(1, 29, d2asign.length, 1).setValues(d2asign);  //NOMBRE INTERNO SEGUN OPC
      SheetDB.getRange(1, 30, d2asign.length, 1).setValues(d4asign);   //NOMBRE LEGAL
      SheetDB.getRange(1, 31, d1asign.length, 1).setValues(d1asign);   //CVE ASIG
      SheetDB.getRange(1, 32, d2asign.length, 2).setValues(d5asign);  //PONDERA
      SheetDB.getRange(1, 34, d2asign.length, 1).setValues(d3asign);  //CUATRI
      SheetDB.getRange(1, 35, d1asign.length, 1).setValues(d6asign);  //OPCION EDU
      
        opEduAnt =opEduActual
        

    
      var opedu = catAcade.getSheetByName("OPCIONES EDUCATIVAS"); 
      var d1opedu = opedu.getRange("A1:A").getValues();
      var d2opedu = opedu.getRange("C1:D").getValues();
      var d3opedu = opedu.getRange("K1:K").getValues();
      SheetDB.getRange("S1:V").clearContent(); 
      SheetDB.getRange(1, 19, d1opedu.length, 1).setValues(d1opedu);
      SheetDB.getRange(1, 20, d2opedu.length, 2).setValues(d2opedu);
      SheetDB.getRange(1, 22, d3opedu.length, 1).setValues(d3opedu);

      //var plant = catAcade.getSheetByName("PLANTELES");
      //var d1plant = plant.getRange("A1:B").getValues();
      var d1plant="CAFETALES"
        SheetDB.getRange("Y1:Z").clearContent();
        SheetDB.getRange(1, 25, d1plant.length, 1).setValue(d1plant);

      /*var asign = catAcade.getSheetByName("ASIGNATURAS");
        var d1asign = asign.getRange("A1:I").getValues();
        SheetDB.getRange("AC1:AK").clearContent(); 
        SheetDB.getRange(1, 29, d1asign.length, 9).setValues(d1asign);
      */

      var grupo = catAcade.getSheetByName("GRUPOS");
        var d1grupo = grupo.getRange("A1:C").getValues();
        var d2grupo = grupo.getRange("E1:E").getValues();
        var d3grupo = grupo.getRange("G1:H").getValues();
        SheetDB.getRange("AQ1:AV").clearContent(); 
        SheetDB.getRange(1, 43, d1grupo.length, 3).setValues(d1grupo);
        SheetDB.getRange(1, 46, d2grupo.length, 1).setValues(d2grupo);
        SheetDB.getRange(1, 47, d3grupo.length, 2).setValues(d3grupo);

      var peredu = catAcade.getSheetByName("PERIODOS EDUCATIVOS");
        var d1peredu = peredu.getRange("A1:N").getValues();
        SheetDB.getRange("AX1:BK").clearContent(); 
        SheetDB.getRange(1, 50, d1peredu.length, 14).setValues(d1peredu);

      var perso = catAcade.getSheetByName("PERSONAL");
        /*var d1perso = perso.getRange("A1:B").getValues();//29-05-23  Solo prof activo
        var d2perso = perso.getRange("E1:E").getValues();
        SheetDB.getRange("BN1:BP").clearContent(); 
        SheetDB.getRange(1, 66, d1perso.length, 2).setValues(d1perso);
        SheetDB.getRange(1, 68, d2perso.length, 1).setValues(d2perso);
        */
        SheetDB.getRange("BN2:BP").clearContent();
        var irakasleak = [perso.getRange("A1:F").getValues()]
          
        var irakasleenMatrizea=[]
        for (var i=0; i<irakasleak[0].length;i++)
        {
          if(irakasleak[0][i][5]=="DOCENTE")
          {
            if(irakasleak[0][i][2]=="ACTIVO")
            {
              irakasleenMatrizea.push([irakasleak[0][i][0],irakasleak[0][i][1],irakasleak[0][i][4]])

            }
          }
        }
        
      SheetDB.getRange(2, 66, irakasleenMatrizea.length, 3).setValues(irakasleenMatrizea);

      var destino = catAcade.getSheetByName("CARPETAS DESTINO ");
        var d1destino = destino.getRange("A1:C").getValues();
        SheetDB.getRange("BS1:BU").clearContent(); 
        SheetDB.getRange(1, 71, d1destino.length, 3).setValues(d1destino);

  //}
}

function bdalum(){
 const catAlumn = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1kzmgvrbkQ7ehUslRxR4ghkpSeI2yT99M3h-Dyo-EEVY/edit#gid=446758492");

var alumn = catAlumn.getSheetByName("ACTIVOS_FORMATEADO");
   var d1alumn = alumn.getRange("A1:J").getValues();
   var d2alumn = alumn.getRange("N1:N").getValues();
   SheetDB.getRange("BZ1:CJ").clearContent(); 
   SheetDB.getRange(1, 78, d1alumn.length, 10).setValues(d1alumn);
   SheetDB.getRange(1, 88, d2alumn.length, 1).setValues(d2alumn); 

}