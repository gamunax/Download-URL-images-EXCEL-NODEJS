const Excel = require('exceljs');
const download = require('image-downloader');
const fs = require('fs');

downloadImage();

async function downloadImage() {
    /*** LEYENDO EL ARCHIVO EXCEL ***/
    const wb = new Excel.Workbook();
    let excelFile = await wb.xlsx.readFile('data.xlsx');
    let ws = excelFile.getWorksheet('data');
    let data = ws.getSheetValues();

    data.forEach(item => {
        if (item[3].hyperlink) {
            /*** DESCARGAR DE IMAGENES ***/
            const path = 'images';
            const destino = item[4].toString();
            const dir = path + '/' + destino;
            if (!fs.existsSync(dir)){
                fs.mkdirSync(dir);
            }
            
            const name = item[2] + '.jpg';
            const options = {
                url: item[3].hyperlink,
                dest: path + '/' + destino + '/' + name
            }
            download.image(options)
                .then(({ filename }) => {
                    console.log('saved to', filename);
                })
                .catch((err) => console.log(err))
        }
    });    
}