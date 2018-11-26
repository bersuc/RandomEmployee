function Sortear() {
  var sheet = SpreadsheetApp.getActiveSheet();
    
  const PlanInicio = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const PlanOpcoes = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1];
  
  //Busca um random entre 3 opções para colocar um funcionário
  var numFunc = Math.round(Math.random() * (3-1) + 1)-1; 
 
  //Busca um random entre 5 opções para colocar uma loja
  var numLoja = Math.round(Math.random() * (4-1) + 1);
  
  //array com nome do funcionario e insere na planilha inicio
  var ListFunc = PlanOpcoes.getRange('E1:E3').getValues();
  PlanInicio.setActiveCell('B5').setValue(ListFunc[numFunc]);
  
    
  //Array das Lojas
  var ListLojas = PlanOpcoes.getRange('B1:B4').getValues();
  PlanInicio.getRange(5, 4).setValue(ListLojas[numLoja]);
  
  // opção 2 - Modelo diferente do funcionário para buscar nome da loja com base no random linha 11
  // var Loja = PlanOpcoes.getRange(numLoja, 2).getValue();
  // PlanInicio.getRange(5, 4).setValue(Loja);


}
