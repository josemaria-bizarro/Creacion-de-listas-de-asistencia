function aukera()
{
var menu = SpreadsheetApp.getUi()
    .createMenu('OPCIONES')
    .addItem('BD', 'ABRIR')
    .addItem('Crea Lista','CREALISTA')
    .addToUi()
}
function ABRIR() {
  bddata();
  bdalum();
}

function CREALISTA(){
 var ListaEspecial = SheetCARATULA.getRange("X5").getValue();

 switch (ListaEspecial){
  case "SI":
  listasEspecial();
  break;
  default:
  GetDataList();
  break;
 }

}