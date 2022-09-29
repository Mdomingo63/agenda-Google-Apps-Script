function doGet(){
    return HtmlService.createTemplateFromFile('web').evaluate().setTitle("Agenda Google Apps Script");

}

    function obtenerDatosHtml(nombre){

       return HtmlService.createHtmlOutputFromFile(nombre).getContent();
    }

       function obtenerContactos(){
             
            let hoja = SpreadsheetApp.openById("1FY5x3sEIB0aKtt2jkaiAjowL_kIYqGiUU8lfj5TPtWo").getActiveSheet();

            let datos = hoja.getDataRange().getValues();
            return datos;

       }