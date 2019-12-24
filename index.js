/*Requerimos la libreria de manipulacion de archivos */
const fs = require('fs');

/*Requerimos  de la libreria para convertir tablas en Excel a una matriz*/
const XLSX = require('xlsx');

/*Para escribir la un objeto a cualquier tipo de archivo */
const util = require('util');

fs.readFile('Planilla.xlsx', (err, arraybuffer) => {
  if (err) throw err;
  console.log(arraybuffer);

  /* Convertimos la data del excel a binaria */
  const data = new Uint8Array(arraybuffer);
  let arr = new Array();
  for(let i = 0; i != data.length; ++i) arr[i] = String.fromCharCode(data[i]);
  const bstr = arr.join("");

  /* Invocamos a la libreria para leer la data binaria */
  const workbook = XLSX.read(bstr, {type:"binary"});

  /* DO SOMETHING WITH workbook HERE */
  const first_sheet_name = workbook.SheetNames[0];
  /* Get worksheet */
  const worksheet = workbook.Sheets[first_sheet_name];
  console.log(XLSX.utils.sheet_to_json(worksheet,{raw:true}));


  /*Reescribiendo data en json*/
 let jsonData = XLSX.utils.sheet_to_json(worksheet,{raw:false});

  fs.writeFile('planilla.json', util.inspect(jsonData, false, 2, false),(err)=>{
    if (err) throw err;
    console.log('Archivo correctamente creado!');
  });
});
