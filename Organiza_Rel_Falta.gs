function formatarPlanilha() {
  var pagina = SpreadsheetApp.getActiveSheet();
  var dados = pagina.getDataRange().getValues();

  // duplica a página atual e define a nova página como ativa
  var paginaOrganizada = pagina.copyTo(pagina.getParent()).setName("Organizada");
  SpreadsheetApp.setActiveSheet(paginaOrganizada);

  // Remove as 5 primeiras linhas da planilha, pois possuem dados que nao interessam
  paginaOrganizada.deleteRows(1, 5);
  dados = paginaOrganizada.getDataRange().getValues();

  // Remove o símbolo de porcentagem da coluna "Percentual de Presença"
  var indiceColunaPercentual = dados[0].indexOf("Percentual de Presença");
  var colunaPercentual = paginaOrganizada.getRange(1, indiceColunaPercentual + 1, paginaOrganizada.getLastRow(), 1);
  colunaPercentual.createTextFinder('%').replaceAllWith('');

  // Altera a fonte, tamanho e cor do texto de toda a página organizada
  paginaOrganizada.getDataRange().setFontFamily('Arial').setFontSize(10)//.setFontColor('#000000');
}