/* 
https://script.google.com/macros/s/AKfycbzccRmC4PsLktVjB1bLEHYp7nscG88eKEiS2KtDPDF8epwfIaAR0yObWHONsse1TfB-/exec 
*/


function doPost(e) {

  const paquete = JSON.parse(e.postData.contents)

  writeEverything(paquete)




return "ok"

}


class SS_ {
  constructor() {
    this.ss = SpreadsheetApp.getActive();
  }
  getSheetByID(gid) {
    const sheet = this.ss.getSheets().filter(sheet => sheet.getSheetId() === gid)[0];
    return sheet;
  }
}

const THIS_SS = new SS_ 

const thisSheet = {
  ss : THIS_SS.ss,
  contrataciones: THIS_SS.getSheetByID(0),
  contratacionesDetalles: THIS_SS.getSheetByID(2094421038),
  ofertasRecibidas: THIS_SS.getSheetByID(303751213),
  condicionesOfertas: THIS_SS.getSheetByID(35697270),
  maestroProveedores: THIS_SS.getSheetByID(413112520),
}

function writeEverything (payload) {
  const paquete = payload
  writeProveedor(paquete.condicionesOfertas)
  const idCont = writeGeneralesContratacion(paquete)
  writeDetalles(paquete.detallesContratacion, idCont)
  writeCondiciones(paquete.condicionesOfertas, idCont)
  writeOfertasRecibidas(paquete.ofertasRecibidas, idCont)

}