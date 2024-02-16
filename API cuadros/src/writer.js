function writeProveedor (condicionesOfertas){
    const provList = thisSheet.maestroProveedores.getDataRange().getValues().shift()


    Object.keys(condicionesOfertas).forEach(key => {

        const provName = quitarAcentos( condicionesOfertas[key].nombreProveedor.replaceAll(".", "").replaceAll(",", "").replaceAll("-","").replaceAll("  "," ").trim().toUpperCase())
        
        const idRow = provList.find(row => row[1] == nomb)

        if(!idRow){
            const newID = provList.map(row => row[0]).max() + 1
            thisSheet.maestroProveedores.appendRow([newID, provName])
    
            condicionesOfertas[key].idProveedor = newID
        }else{
            condicionesOfertas[key].idProveedor = idRow[0]
        }

        
    })

}
function quitarAcentos(cadena){
	const acentos = {'á':'a','é':'e','í':'i','ó':'o','ú':'u','Á':'A','É':'E','Í':'I','Ó':'O','Ú':'U'};
	return cadena.split('').map( letra => acentos[letra] || letra).join('').toString();	
}
function writeGeneralesContratacion(paquete){
  const grals = paquete.generalesContratacion

  const timestamp = new Date()

  const id =  Math.random().toString(36).substr(2, 9);

    const row = [
        timestamp,
        id,
        grals.tipoProc + " " + grals.numProc + "/" + grals.anoProc,
        grals.tipoProc,
        grals.numProc,
        grals.anoProc,
        grals.organismoProc,
        grals.categoriaProc,
        grals.objProc,
        paquete.detallesContratacion.length,
        Object.keys(paquete.condicionesOfertas).length,
        Number(grals.presupProc)
    ]
    thisSheet.contrataciones.appendRow(row)
    return id

}
function   writeDetalles(detallesContratacion, idCont){
    const rowToInsert = thisSheet.contratacionesDetalles.getLastRow() + 1
    const rows = detallesContratacion.map(detalle => {
        return [
            idCont,
            detalle[0],
            detalle[1],
            detalle[2],
            detalle[3],
            detalle[4],
        ]
    }
    )
    thisSheet.contratacionesDetalles.getRange(rowToInsert, 1, rows.length, rows[0].length).setValues(rows)

}
function   writeCondiciones(condicionesOfertas, idCont){
    const rowToInsert = thisSheet.condicionesOfertas.getLastRow() + 1
    const rows = []

    Object.keys(condicionesOfertas).forEach(key => {
        rows.push([
            idCont,
            key,
            condicionesOfertas[key].idProveedor,
            condicionesOfertas[key].formaEntrega,
            condicionesOfertas[key].formaPago,
            condicionesOfertas[key].mantenimientoOferta,
        ])
    })
    thisSheet.condicionesOfertas.getRange(rowToInsert, 1, rows.length, rows[0].length).setValues(rows)
}

function   writeOfertasRecibidas(ofertasRecibidas, idCont, condicionesOfertas){
    const rowToInsert = thisSheet.ofertasRecibidas.getLastRow() + 1
    const rows = []

    Object.keys(ofertasRecibidas).forEach(renglon => {
        Object.keys(ofertasRecibidas[renglon]).forEach(ordenMerito => {
            const objOferta = ofertasRecibidas[renglon][ordenMerito]
            const row = [
                idCont,
                renglon,
                ordenMerito,
                objOferta.nAlt,
                objOferta.nProv,
                condicionesOfertas[objOferta.nProv].idProveedor,
                Number(objOferta.qOfert),
                Number(objOferta.pUnit),
                Number(objOferta.qOfert * objOferta.pUnit),
                objOferta.observacion
            ]
            rows.push(row)
        })
    })
    thisSheet.ofertasRecibidas.getRange(rowToInsert, 1, rows.length, rows[0].length).setValues(rows) 
}