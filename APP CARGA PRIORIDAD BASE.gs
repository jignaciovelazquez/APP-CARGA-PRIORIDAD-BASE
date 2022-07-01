function doGet() {
  var html = HtmlService.createTemplateFromFile('Index.html').evaluate()
  .setTitle("Carga Prioridad")
  .setFaviconUrl("https://m1.paperblog.com/i/232/2324416/boton-subir-jquery-que-aparece-desaparce-R-IWQjNW.png");
  return html

}

function buscarID(id){
  const libro = SpreadsheetApp.openById("1iyba6EH-qooC6mA3jMy1NDFpf42bwcx1Uip-tGwsvT4");
  const Hoja = libro.getSheetByName("Online");  
  var UltimaFila = Hoja.getLastRow();
  

  var duplicado = 0;
  var fila = 0;

  for(i=1;i<=UltimaFila;i++){
    if (Hoja.getRange(i,3).getValue()==id){
      duplicado = 1;
      fila = i;
      //Hoja.getRange(i,2).setBackground("red");
    }
  }
  return [duplicado,fila];
}


function Escribir(hora,dispatcher,id,nodo,direccion,diseno,gestion,contratista,tecnico,fecha,obs){
  const libro = SpreadsheetApp.openById("1iyba6EH-qooC6mA3jMy1NDFpf42bwcx1Uip-tGwsvT4");
  const Hoja = libro.getSheetByName("Online");  
  var UltimaFila = Hoja.getLastRow();
  const Hoja2 = libro.getSheetByName("GESTIONES");  
  var UltimaFila2 = Hoja.getLastRow();

  var gestionduplicada = 0;

  for(i=1;i<=UltimaFila2;i++){
    if (Hoja2.getRange(i,1).getValue()==id){
      gestionduplicada = 1;
    }
  }

  if (gestionduplicada==0){
    Hoja.getRange(UltimaFila+1,1).setValue(hora);
    Hoja.getRange(UltimaFila+1,2).setValue(dispatcher);
    Hoja.getRange(UltimaFila+1,3).setValue(id);
    Hoja.getRange(UltimaFila+1,4).setValue(nodo);
    Hoja.getRange(UltimaFila+1,5).setValue(direccion);
    Hoja.getRange(UltimaFila+1,6).setValue(diseno);
    Hoja.getRange(UltimaFila+1,7).setValue(gestion);
    Hoja.getRange(UltimaFila+1,8).setValue(contratista);
    Hoja.getRange(UltimaFila+1,9).setValue(tecnico);
    Hoja.getRange(UltimaFila+1,10).setValue(fecha);
    Hoja.getRange(UltimaFila+1,11).setValue(obs);

  }else{
    Hoja.getRange(UltimaFila+1,1).setValue(hora);
    Hoja.getRange(UltimaFila+1,2).setValue(dispatcher);
    Hoja.getRange(UltimaFila+1,3).setValue(id);
    Hoja.getRange(UltimaFila+1,4).setValue(nodo);
    Hoja.getRange(UltimaFila+1,4).setBackground("red");
    Hoja.getRange(UltimaFila+1,5).setValue(direccion);
    Hoja.getRange(UltimaFila+1,5).setBackground("red");
    Hoja.getRange(UltimaFila+1,6).setValue(diseno);
    Hoja.getRange(UltimaFila+1,7).setValue(gestion);
    Hoja.getRange(UltimaFila+1,8).setValue(contratista);
    Hoja.getRange(UltimaFila+1,9).setValue(tecnico);
    Hoja.getRange(UltimaFila+1,10).setValue(fecha);
    Hoja.getRange(UltimaFila+1,11).setValue(obs);
  }


    return
}