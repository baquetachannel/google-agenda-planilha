/**
 * Copia os eventos da Agenda Google padrão dentro do intervalo de dias [inicio, fim] para a página de uma planilha.
 * 
 * @param pagina A página da planilha onde serão escritos os eventos.
 * @param {Date} inicio A data início do intervalo.
 * @param {Date} fim A data final do intervalo.
 */
function atualizarPlanilha(pagina, inicio, fim) {
  var agenda = CalendarApp.getDefaultCalendar();
  pagina.getRange('A1:F1').setValues([['ID Evento', 'Título', 'Descrição', 'Data Início', 'Data Fim', 'ID Cor']]);

  var eventos = agenda.getEvents(inicio, fim);

  for (var i = 0; i < eventos.length; i++) {
    var evento = eventos[i];

    var idEvento = evento.getId();
    var titulo = evento.getTitle();
    var descricao = evento.getDescription();
    var dataInicio = evento.getStartTime();
    var dataFim = evento.getEndTime();
    var idCor = evento.getColor();

    Logger.log(idEvento + ' | ' + titulo + ' | ' + descricao + ' | ' + dataInicio + ' | ' + dataFim + ' | ' + idCor );

    pagina.appendRow([idEvento, titulo, descricao, dataInicio, dataFim, idCor]);
  }
}


/**
 * Copia os eventos de uma Agenda Google dentro dos últimos 4 dias até a data de EXECUÇÃO, 
 * ou seja, aproximadamente 5 dias, dependendo da data de execução.
 * Também cria uma página nova para salvar os eventos na planilha no formato: DD/MM - DD/MM
 *  
 * Exemplo: se for executado dia 2024-01-05 às 22h, criará uma página "01/01 - 05/01" na planilha
 * correspondente e copiará todos os eventos a partir de 0h do dia 2024-01-01 até às 22h do dia 
 * 2024-01-05 pra página criada.
 * 
 */
function atualizarPlanilhaSemanal() {
  var fim = new Date();
  var inicio = new Date();
  inicio.setDate(fim.getDate() - 4);  // Subtrai 4 dias da data de fim
  inicio.setHours(0, 0, 0);  // 0h 

  var inicioFmt = inicio.toLocaleDateString('pt-BR', {day: '2-digit', month: '2-digit'});
  var fimFmt = fim.toLocaleDateString('pt-BR', {day: '2-digit', month: '2-digit'});

  // Cria página
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  planilha.insertSheet(`${inicioFmt} - ${fimFmt}`);  // Formata pra DD/MM - DD/MM
  var pagina = planilha.getActiveSheet();
  
  atualizarPlanilha(pagina, inicio, fim);
}

