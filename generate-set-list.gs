// @NotOnlyCurrentDoc
// 
// all thanks to Toncoso who posted this on stack overflow
// https://stackoverflow.com/questions/22362504/use-google-script-to-separate-a-google-doc
// Updated 6/1/2018 for driveapp
//
// Lucas 9/8/21: Atualizando para dividir páginas específicas
// 
//  "oauthScopes": [
//    "https://www.googleapis.com/auth/drive",
//    "https://www.googleapis.com/auth/documents",
//    "https://www.googleapis.com/auth/spreadsheets.currentonly",
//    "https://www.googleapis.com/auth/script.external_request"],
// 
// Truque para rodar sempre como o owner:
// https://tanaikech.github.io/2020/11/05/user-runs-script-for-range-protected-by-owner-using-google-apps-script/

function GerarCifras() {
  var PB  = DocumentApp.ElementType.PAGE_BREAK;
  SpreadsheetApp.getActive().setActiveSheet(SpreadsheetApp.getActive().getSheetByName("Programa"));
  Logger.log("Gerando...")
  SpreadsheetApp.getActiveSheet().getRange(ProgramaStatusCell).setValue("Gerando...")
  // Abre pasta de destino
  Logger.log("Folder id = " + SpreadsheetApp.getActiveSheet().getRange("proximo_culto").getValue().match(/[-\w]{25,}/))
  var folder_proximo_culto = DriveApp.getFolderById(SpreadsheetApp.getActiveSheet().getRange("proximo_culto").getValue().match(/[-\w]{25,}/));
  SpreadsheetApp.getActiveSheet().getRange(ProgramaStatusCell).setValue("Folder ok...")
  Logger.log("Folder name = " + folder_proximo_culto)
  if (folder_proximo_culto == null) {
    SpreadsheetApp.getActiveSheet().getRange(ProgramaStatusCell).setValue("Folder não encontrada!")
    Logger.log("Folder não encontrada!")
    return false;
  }
  // Abre doc com todas as cifras
  SpreadsheetApp.getActiveSheet().getRange(ProgramaStatusCell).setValue("Abrindo...")
  var doc_todas_as_cifras = DocumentApp.openByUrl(SpreadsheetApp.getActiveSheet().getRange("todas_as_cifras").getValue());
  var par = doc_todas_as_cifras.getBody().getParagraphs();
  // Cria doc do proximo culto
  SpreadsheetApp.getActiveSheet().getRange(ProgramaStatusCell).setValue("Criando...")
  Logger.log("Criando doc_proximo_culto...")
  //var doc_proximo_culto = DocumentApp.create("Cifras proximo domingo");
  var doc_proximo_culto_file = DriveApp.getFileById(doc_todas_as_cifras.getId()).makeCopy(Utilities.formatDate(SpreadsheetApp.getActiveSheet().getRange("A12").getValue(),"GMT-3","yyyy-MM-dd") + " Cifras", folder_proximo_culto)
  var doc_proximo_culto = DocumentApp.openById(doc_proximo_culto_file.getId())
  Logger.log("Filename: " + doc_proximo_culto.getName())
  // Limpa o documento para começar a inserir as músicas do próximo culto
  doc_proximo_culto.getBody().setText("")
  // Pega os números das músicas da planilha
  var lista_de_musicas = SpreadsheetApp.getActiveSheet().getRange("B12:B23").getValues(); // TODO gerar do dia selecionado
  SpreadsheetApp.getActiveSheet().getRange(ProgramaStatusCell).setValue("Adicionando...")
  // Varre a lista de musicas do proximo domingo
  for (var musica = 0; musica < lista_de_musicas.length; musica++) {
    if (lista_de_musicas[musica] == "") { continue }
    // Varre as paginas até chegar na musica
    SpreadsheetApp.getActiveSheet().getRange(ProgramaStatusCell).setValue("Música " + lista_de_musicas[musica] + "...")
    Logger.log("Procurando Musica = " + lista_de_musicas[musica])
    var page = 1;
    for (var i = 0; i < par.length; i++) {
      if (page == lista_de_musicas[musica]) {
        doc_proximo_culto.getBody().appendParagraph(par[i].copy());
        Logger.log("Adding par = " + i)
      }
      // Procura por um Page Break em cada Paragrafo, quando encontra, incrementa a página
      for (var j = 0; j < par[i].getNumChildren(); j++) {
        if (par[i].getChild(j).getType() == PB) {
          page++;
          break;
        }
      }
      // Musica terminada, vai para proxima
      if (page > lista_de_musicas[musica]) {
        Logger.log("Musica terminada = " + lista_de_musicas[musica])
        break;
      }
    }
  }
  SpreadsheetApp.getActiveSheet().getRange(ProgramaStatusCell).setValue("Salvando...")
  Logger.log("Salvando...")
  doc_proximo_culto.saveAndClose();
  doc_todas_as_cifras.saveAndClose();
  SpreadsheetApp.getActiveSheet().getRange(ProgramaStatusCell).setValue("Pronto")
  Logger.log("Pronto")
  SpreadsheetApp.getUi().alert("\"" + Utilities.formatDate(SpreadsheetApp.getActiveSheet().getRange("A12").getValue(),"GMT-3","yyyy-MM-dd") + " Cifras\" gerado com sucesso na pasta \"Próximo Culto\"");
}
