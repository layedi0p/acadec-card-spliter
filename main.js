const { app, BrowserWindow, dialog } = require('electron')
const xlsx = require('xlsx');
const fs = require('fs');
const https = require('https');
const path = require('path');
const converter = require('json-2-csv');

const createWindow = () => {
    const win = new BrowserWindow({
        width: 450,
        height: 350,
    })
    win.loadFile('index.html')
    dialog.showOpenDialog(win, {
        properties: ['openFile'],
        filters: [
            { name: 'Fichiers Excel', extensions: ['xlsx', 'xls'] }
        ]
    }).then(result => {
        if (!result.canceled) {
            const filePath = result.filePaths[0];
            try {
                const jsonData = excelToJson(filePath);
                console.log(jsonData);
                //select save dir
                dialog.showOpenDialog(win, {
                    properties: ['openDirectory', 'createDirectory']
                }).then(async result => {
                    if (!result.canceled) {
                        const downloadDir = result.filePaths[0];
                        fs.mkdirSync(path.join(downloadDir, 'photos'), { recursive: true });
                        fs.mkdirSync(path.join(downloadDir, 'qrcodes'), { recursive: true });
                        let counter = 0;
                        for (const row of jsonData) {
                            counter++;
                            const imageName = `image${counter}.jpg`;
                            const qrcodeName = `qrcode${counter}.jpg`;
                            if (row.photo) {
                                await downloadFile(row.photo, path.join(downloadDir, 'photos'), imageName)
                                    .then(filePath => {
                                        console.log(`Photo téléchargée: ${filePath}`);
                                    }).catch(err => {
                                        console.error(err.message);
                                    });
                            }
                            if (row.qrcode) {
                                await downloadFile(row.qrcode, path.join(downloadDir, 'qrcodes'), qrcodeName)
                                    .then(filePath => {
                                        console.log(`QRCode téléchargé: ${filePath}`);
                                    }).catch(err => {
                                        console.error(err.message);
                                    });
                            }
                            row.photo = path.join(downloadDir, 'photos', imageName);
                            row.qrcode = path.join(downloadDir, 'qrcodes', qrcodeName);
                        }


                        const treattedData = [];
                        const chunkedData = chunkArray(jsonData, 10);
                        for (const data of chunkedData) {
                            treattedData.push(data.reduce((acc, curr, index) => {
                                const response = {
                                    ...acc,
                                }
                                Object
                                    .keys(curr)
                                    .forEach(key => {
                                        let header = key + (index + 1);
                                        if (['photo', 'qrcode'].includes(key)) {
                                            header = '@' + header;
                                        } else if (key === 'nomComplet') {
                                            let name = curr[key];
                                            if(name.length >= 18 ) {
                                                const names = name.split(' ');
                                                name = names[0];
                                                for(let i = 1; i < names.length - 1; i++) {
                                                   name += ' ' + names[i].charAt(0) + '.';
                                                }
                                                name += ' ' + names[names.length - 1];
                                            }
                                            curr[key] = name;
                                        }
                                        response[header] = curr[key];
                                    });
                                return response;
                            }, {}));
                        }
                        //const book = xlsx.utils.book_new();
                        //xlsx.utils.book_append_sheet(book, xlsx.utils.json_to_sheet(treattedData), 'Sheet1');
                        //write xlsx to csv
                        //const csv = xlsx.utils.sheet_to_csv(book.Sheets[book.SheetNames[0]]);
                        const csv = converter.json2csv(treattedData, {
                            emptyFieldValue: '',
                            unwindArrays: true,
                            unwindArraysSeparator: ',',
                            flatten: true
                        });
                        fs.writeFileSync(path.join(downloadDir, 'output.csv'), csv, 'utf-8');
                        process.exit(0);
                    }
                }).catch(err => {
                    console.error(err);
                })


            } catch (error) {
                console.error(error.message);
            }
        }
    }).catch(err => {
        console.error(err);
    })
}

app.whenReady().then(() => {
    createWindow()
    app.on('activate', () => {
        if (BrowserWindow.getAllWindows().length === 0) createWindow()
    })
})

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit()
})


/**
 * Convertit la première feuille d'un fichier Excel en JSON.
 * @param {string} filePath - Chemin du fichier Excel.
 * @returns {Object[]} - Données de la première feuille sous forme de tableau JSON.
 */
function excelToJson(filePath) {
    // Vérifier si le fichier existe
    if (!fs.existsSync(filePath)) {
        throw new Error(`Le fichier n'existe pas: ${filePath}`);
    }

    // Charger le fichier Excel
    const workbook = xlsx.readFile(filePath);

    // Obtenir le nom de la première feuille
    const firstSheetName = workbook.SheetNames[0];

    // Convertir la feuille en JSON
    const jsonData = xlsx.utils.sheet_to_json(workbook.Sheets[firstSheetName], { raw: false });

    return jsonData;
}


/**
 * Télécharge un fichier depuis une URL et l'enregistre dans un répertoire donné.
 * @param {string} fileUrl - URL du fichier à télécharger.
 * @param {string} downloadDir - Répertoire de destination.
 * @param {string} fileName - Nom sous lequel enregistrer le fichier.
 * @returns {Promise<string>} - Chemin du fichier téléchargé.
 */
function downloadFile(fileUrl, downloadDir, fileName) {
    console.log('downloading file', fileUrl);
    return new Promise((resolve, reject) => {
        // Vérifier si le répertoire existe, sinon le créer
        if (!fs.existsSync(downloadDir)) {
            fs.mkdirSync(downloadDir, { recursive: true });
        }

        const filePath = path.join(downloadDir, fileName);
        const file = fs.createWriteStream(filePath);

        https.get(fileUrl, (response) => {
            if (response.statusCode !== 200) {
                return reject(new Error(`Échec du téléchargement, code HTTP: ${response.statusCode}`));
            }
            response.pipe(file);
            file.on('finish', () => {
                file.close(() => resolve(filePath));
            });
        }).on('error', (err) => {
            fs.unlink(filePath, () => reject(err));
        });
    });
}

/**
 * Divise un tableau en sous-tableaux de taille égale.
 * @param {Array} array - Le tableau à diviser.
 * @param {number} chunkSize - La taille de chaque sous-tableau.
 * @returns {Array[]} - Un tableau contenant les sous-tableaux.
 */
function chunkArray(array, chunkSize) {
    if (chunkSize <= 0) throw new Error("La taille du chunk doit être supérieure à 0");
    return Array.from({ length: Math.ceil(array.length / chunkSize) }, (_, i) =>
        array.slice(i * chunkSize, i * chunkSize + chunkSize)
    );
}