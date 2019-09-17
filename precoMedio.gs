//Calcular preço médio dos ativos
function PRECOMEDIO(ativo){
  var spreadsheet = SpreadsheetApp.getActive().getSheetByName("Log");
  spreadsheet.getRange('A1:I1').activate();
  var data = spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate().getValues();
  data = data.sort();
  
  var qtdTotal = 0;
  for(var key_row in data){
    //Logger.log(data[key_row]);
    if(ativo == data[key_row][3]){
      Logger.log(data[key_row][3])
      qtdTotal += data[key_row][4];
    }
  }
  Logger.log("Qtd Total: ", qtdTotal);
  var dataReverse = data.reverse();
  
  var precoTotal = 0;
  var qtdTotalAux = qtdTotal;
  
  for(var key_row in dataReverse){
    if(qtdTotalAux < 0){
     break; 
    }
    var row = dataReverse[key_row]; 
    if(ativo == row[3] && row[2] == "C"){
      var qtd = row[4];
      var precoUnid = row[5];
      if(qtdTotalAux >= qtd){
        precoTotal += qtd * precoUnid;
        qtdTotalAux -= qtd;
      } else{
        precoTotal += qtdTotalAux * precoUnid;
        qtdTotalAux = 0;
      }
    }
  }
  Logger.log('Preco Total:',precoTotal);
  
  var precoMedio = 0;
  qtdTotal == 0 ? precoMedio = 0 : precoMedio = precoTotal/qtdTotal;
  return precoMedio;
}