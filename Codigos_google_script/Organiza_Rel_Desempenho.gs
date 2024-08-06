function divideColunaNota() {
  var pagina = SpreadsheetApp.getActiveSheet();
  var range = pagina.getDataRange();
  var valores = range.getValues();

  // duplica a página atual e define a nova página como ativa
  var paginaOrganizada = pagina.copyTo(pagina.getParent()).setName("Organizada");
  SpreadsheetApp.setActiveSheet(paginaOrganizada);

  var header = valores.shift();
  var notaIndex = header.indexOf("Nota");

  var notas = {};
  var media = [];

  for (var i = 0; i < valores.length; i++) {
    var linha = valores[i];
    var notaValor = linha[notaIndex];
    var notaSplit = notaValor.split(";");

    // Rodar nas notas da linha e salvar os valores
    for (var j = 0; j < notaSplit.length; j++) {
      var nota = notaSplit[j];
      var notaInfo = nota.split(":");
      var notaNome = notaInfo[0];
      var notaValor = notaInfo[1];

      // Verificar se é uma nota e colocar seu respectivo valor
      if (notaNome.indexOf("A") == 0) {
        if (!notas[notaNome]) {
          notas[notaNome] = [];
        }
        notas[notaNome][i] = notaValor; 
      } else if (notaNome.indexOf("Media") == 0) {
        media[i] = notaValor;
      }
    }
  }

  for (var notaNome in notas) {
    header.push(notaNome);
    var colIndex = header.length - 1;
    for (var i = 0; i < valores.length; i++) {
      var linha = valores[i];
      if (!notas[notaNome][i]) {
        linha[colIndex] = "-";
      } else {
        linha[colIndex] = notas[notaNome][i];
      }
    }
  }

  header.push("Media");
  var colIndex = header.length - 1;
  for (var i = 0; i < valores.length; i++) {
    var linha = valores[i];
    if (!media[i]) {
      linha[colIndex] = "-";
    } else {
      linha[colIndex] = media[i];
    }
  }

  // Pergunta ao usuário o valor numérico para a coluna "ANO"
  var anoValue = Browser.inputBox("Digite o valor numérico para a coluna ANO:");

  // Adiciona a coluna "ANO" com o valor fornecido em todas as linhas
  header.push("ANO");
  var anoIndex = header.length - 1;
  for (var i = 0; i < valores.length; i++) {
    valores[i][anoIndex] = anoValue;
  }

  // Altera a fonte, tamanho e cor do texto de toda a página organizada
  paginaOrganizada.getDataRange().setFontFamily('Arial').setFontSize(10).setFontColor('#000000');

  paginaOrganizada.clearContents();
  paginaOrganizada.getRange(1, 1, 1, header.length).setValues([header]);
  paginaOrganizada.getRange(2, 1, valores.length, header.length).setValues(valores);

  paginaOrganizada.deleteColumn(notaIndex + 1);
}
