const fs = require('fs');
const path = require('path');
const csv = require('csv-parser');
const ExcelJS = require('exceljs');
const Koa = require('koa');
const Router = require('koa-router');
const serve = require('koa-static');
const multer = require('@koa/multer');
const packageJson = require('../package.json');

const app = new Koa();
const router = new Router();
const upload = multer({ dest: './src/public/uploads/' });

multer({ dest: './src/public/input/' });
multer({ dest: './src/public/output/' });

async function processCSV(fileName) {
    const csvFilePath = path.join(__dirname, './public/input/' + fileName);

    const results = [];

    fs.createReadStream(csvFilePath)
        .pipe(csv({ separator: ';' }))
        .on('data', (data) => {
            const existingEntry = results.find(entry => entry.tienda === data.tienda && entry.semana === data.semana && entry.nombre === data.nombre);
            if (existingEntry) {
                // Update the existing entry if needed
            } else {
                results.push(data);
            }
        })
        .on('end', () => {
            const employeeCountByWeek = results.reduce((acc, entry) => {
                if (!acc[entry.semana]) {
                acc[entry.semana] = { count: 0, totalHours: 0 };
                }
                acc[entry.semana].count++;
                acc[entry.semana].totalHours += parseFloat(entry.horas);
                return acc;
            }, {});

            const averageHoursByWeek = Object.keys(employeeCountByWeek).reduce((acc, week) => {
                acc[week] = (employeeCountByWeek[week].totalHours / employeeCountByWeek[week].count).toFixed(2);
                return acc;
            }, {});

            const combinedData = Object.keys(employeeCountByWeek).map(week => ({
                semana: week,
                empleados: employeeCountByWeek[week].count,
                horasMedias: averageHoursByWeek[week],
                horasTotales: employeeCountByWeek[week].totalHours
            }));

            // Crea una lista de los empleados por semana con el número de horas totales
            const employeeListByWeek = results.reduce((acc, entry) => {
                if (!acc[entry.semana]) {
                    acc[entry.semana] = [];
                }
                acc[entry.semana].push({
                    nombre: entry.nombre,
                    horas: parseFloat(entry.horas)
                });
                return acc;
            }, {});

            // Haz una lista por cada empleado de las horas trabajadas por semana
            const weeks = [...new Set(results.map(entry => entry.semana))];
            const hoursByEmployee = results.reduce((acc, entry) => {
                if (!acc[entry.nombre]) {
                acc[entry.nombre] = {};
                weeks.forEach(week => {
                    acc[entry.nombre][week] = 0;
                });
                }
                acc[entry.nombre][entry.semana] += parseFloat(entry.horas);
                return acc;
            }, {});

            //console.log('Horas trabajadas por empleado por semana:', hoursByEmployee);

            // Haz un ranking de las horas totales trabajadas por empleado
            const employeeRanking = results.reduce((acc, entry) => {
                if (!acc[entry.nombre]) {
                    acc[entry.nombre] = 0;
                }
                acc[entry.nombre] += parseFloat(entry.horas);
                return acc;
            }, {});

            const sortedEmployeeRanking = Object.keys(employeeRanking)
                .map(name => ({ nombre: name, horasTotales: employeeRanking[name] }))
                .sort((a, b) => b.horasTotales - a.horasTotales);

            // Añade al ranking anterior el número de semanas del año trabajadas
            sortedEmployeeRanking.forEach(employee => {
                employee.semanas = results.reduce((acc, entry) => {
                    if (entry.nombre === employee.nombre && !acc.includes(entry.semana)) {
                        acc.push(entry.semana);
                    }
                    return acc;
                }, []).length;
            });

            // Añade la media y mediana de horas trabajadas por semana de cada empleado al ranking
            const median = (arr) => {
                const sorted = arr.slice().sort((a, b) => a - b);
                const mid = Math.floor(sorted.length / 2);
                return sorted.length % 2 !== 0 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
            };

            sortedEmployeeRanking.forEach(employee => {
                const weeklyHours = results
                    .filter(entry => entry.nombre === employee.nombre)
                    .map(entry => parseFloat(entry.horas));

                const totalWeeks = weeklyHours.length;
                const totalHours = weeklyHours.reduce((acc, hours) => acc + hours, 0);
                const averageHours = (totalHours / totalWeeks).toFixed(2);
                const medianHours = median(weeklyHours).toFixed(2);

                employee.mediaHoras = averageHours;
                employee.medianaHoras = medianHours;
            });

            // Añade la varianza y desviación estandar de horas por semana de cada empleado
            const variance = (arr) => {
                const mean = arr.reduce((acc, val) => acc + val, 0) / arr.length;
                return arr.reduce((acc, val) => acc + Math.pow(val - mean, 2), 0) / arr.length;
            };

            const standardDeviation = (arr) => {
                return Math.sqrt(variance(arr));
            };

            sortedEmployeeRanking.forEach(employee => {
                const weeklyHours = results
                .filter(entry => entry.nombre === employee.nombre)
                .map(entry => parseFloat(entry.horas))
                .filter(hours => hours !== 0); // Eliminate 0 values

                if (weeklyHours.length > 0) {
                    employee.varianzaHoras = variance(weeklyHours).toFixed(2);
                    employee.desviacionEstandarHoras = standardDeviation(weeklyHours).toFixed(2);
                } else {
                    employee.varianzaHoras = '0.00';
                    employee.desviacionEstandarHoras = '0.00';
                }
            });

            // Exporta en un xlsx la lista de horas semanales de la tienda
            const workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet('Reporte General');
            const employeeWorksheet = workbook.addWorksheet('Semana y horas por empleado');
            const rankingWorksheet = workbook.addWorksheet('Ranking empleados');
            const hoursByEmployeeWorksheet = workbook.addWorksheet('Horas por empleado y semana');

            hoursByEmployeeWorksheet.columns = [
                { header: 'Nombre', key: 'nombre', width: 35 },
                ...weeks.map(week => ({ header: `Semana ${week}`, key: `semana_${week}`, width: 10 }))
            ];

            const sortedEmployees = Object.keys(hoursByEmployee).sort((a, b) => {
                const totalHoursA = Object.values(hoursByEmployee[a]).reduce((acc, hours) => acc + hours, 0);
                const totalHoursB = Object.values(hoursByEmployee[b]).reduce((acc, hours) => acc + hours, 0);
                return totalHoursB - totalHoursA;
            });

            sortedEmployees.forEach(employee => {
                const row = { nombre: employee };
                weeks.forEach(week => {
                    row[`semana_${week}`] = hoursByEmployee[employee][week];
                });
                hoursByEmployeeWorksheet.addRow(row);
            });

            rankingWorksheet.columns = [
                { header: 'Nombre', key: 'nombre', width: 35 },
                { header: 'Horas Totales', key: 'horasTotales', width: 15 },
                { header: 'Semanas con turno', key: 'semanas', width: 20 },
                { header: 'Media de horas', key: 'mediaHoras', width: 15 },
                { header: 'Mediana de horas', key: 'medianaHoras', width: 16 },
                { header: 'Desviación estandar de los turnos', key: 'desviacionEstandarHoras', width: 30 }
            ];

            sortedEmployeeRanking.forEach(employee => {
                rankingWorksheet.addRow({
                    nombre: employee.nombre,
                    horasTotales: employee.horasTotales,
                    semanas: employee.semanas,
                    mediaHoras: employee.mediaHoras,
                    medianaHoras: employee.medianaHoras,
                    desviacionEstandarHoras: employee.desviacionEstandarHoras
                });
            });

            rankingWorksheet.eachRow((row, rowNumber) => {
                if (rowNumber > 1) { // Skip header row
                    row.getCell('mediaHoras').value = parseFloat(row.getCell('mediaHoras').value);
                    row.getCell('medianaHoras').value = parseFloat(row.getCell('medianaHoras').value);
                    row.getCell('desviacionEstandarHoras').value = parseFloat(row.getCell('desviacionEstandarHoras').value);
                }
            });

            employeeWorksheet.columns = [
                { header: 'Semana', key: 'semana', width: 10 },
                { header: 'Nombre', key: 'nombre', width: 35 },
                { header: 'Horas', key: 'horas', width: 10 }
            ];

            Object.keys(employeeListByWeek).forEach(week => {
                employeeListByWeek[week].forEach(employee => {
                    employeeWorksheet.addRow({
                        semana: week,
                        nombre: employee.nombre,
                        horas: employee.horas
                    });
                });
            });

            // Combine cells with the same week value
            let startRow = 2; // Start from the first data row
            let endRow = startRow;
            let currentWeek = employeeWorksheet.getCell(`A${startRow}`).value;

            for (let i = startRow + 1; i <= employeeWorksheet.rowCount; i++) {
                const cellValue = employeeWorksheet.getCell(`A${i}`).value;
                if (cellValue === currentWeek) {
                    endRow = i;
                } else {
                    if (startRow !== endRow) {
                        employeeWorksheet.mergeCells(`A${startRow}:A${endRow}`);
                    }
                    startRow = i;
                    endRow = i;
                    currentWeek = cellValue;
                }
            }

            // Merge the last set of cells if needed
            if (startRow !== endRow) {
                employeeWorksheet.mergeCells(`A${startRow}:A${endRow}`);
            }

            worksheet.columns = [
                { header: 'Semana', key: 'semana', width: 10 },
                { header: 'Empleados', key: 'empleados', width: 13 },
                { header: 'Horas Totales', key: 'horasTotales', width: 15 },
                { header: 'Media de horas por empleado', key: 'horasMedias', width: 30 }
            ];

            combinedData.forEach(data => {
                worksheet.addRow({
                    semana: data.semana,
                    empleados: data.empleados,
                    horasTotales: data.horasTotales,
                    horasMedias: data.horasMedias
                });
            });

            worksheet.eachRow((row, rowNumber) => {
                if (rowNumber > 1) { // Skip header row
                    row.getCell('C').value = parseFloat(row.getCell('C').value);
                    row.getCell('D').value = parseFloat(row.getCell('D').value);
                }
            });
            
            const outputFilePath = path.join(__dirname, './public/output/', `report_${fileName.replace('.csv', '')}.xlsx`);

            const fileNew = workbook.xlsx.writeFile(outputFilePath);

            fs.unlink(csvFilePath, (err) => {
                if (err) {
                    console.error('Error al eliminar el archivo CSV:', err);
                } else {
                    console.log('Archivo CSV eliminado correctamente.');
                }
            });

            if (fileNew) {
                return fileNew;
            } else {
                return false;
            }
            
        });
}

router.get('/', async (ctx) => {
    ctx.body = `
        <html>
            <head>
                <title>Shop Reporter</title>
                <style>
                    body {
                        background-color: #f8f8f8;
                        color: #333333;
                        text-align: center;
                        font-family: 'Verdana', sans-serif;
                        margin: 0;
                        padding: 0;
                        display: flex;
                        flex-direction: column;
                        justify-content: center;
                        align-items: center;
                        height: 100vh;
                    }
                    a {
                        color: #ff6f61;
                        text-decoration: none;
                        font-weight: bold;
                    }
                    a:hover {
                        text-decoration: underline;
                    }
                    img {
                        width: 150px;
                        margin-top: 20px;
                        margin-bottom: 30px;
                    }
                    .container {
                        display: flex;
                        flex-direction: column;
                        padding: 20px;
                        width: 80%;
                        max-width: 1200px;
                    }
                    .row {
                        display: flex;
                        width: 100%;
                        justify-content: center;
                    }
                    .title {
                        font-size: 2.5em;
                        font-weight: bold;
                        color: #333333;
                        text-transform: uppercase;
                        letter-spacing: 2px;
                        margin: 0;
                        padding: 10px 20px;
                        border-bottom: 2px solid #ff6f61;
                    }
                    .column {
                        flex: 1;
                        padding: 20px;
                        border: 1px solid #333333;
                        border-radius: 10px;
                        margin: 10px;
                    }
                    h1 {
                        margin-top: 20px;
                        font-size: 2em;
                    }
                    p {
                        font-size: 1em;
                        line-height: 1.5;
                    }
                    .button {
                        display: inline-block;
                        padding: 10px 20px;
                        margin-top: 20px;
                        background-color: #ff6f61;
                        color: #ffffff;
                        border: none;
                        border-radius: 5px;
                        cursor: pointer;
                        font-size: 1em;
                    }
                    .button:hover {
                        background-color: #e65c50;
                    }
                    input[type="file"] {
                        margin-top: 20px;
                        padding: 10px;
                        border: 1px solid #ff6f61;
                        border-radius: 5px;
                        background-color: #ffffff;
                        color: #333333;
                        cursor: pointer;
                        font-size: 1em;
                        transition: background-color 0.3s ease;
                    }
                    input[type="file"]:hover {
                        background-color: #ffefef;
                    }
                    button[type="submit"] {
                        margin-top: 20px;
                        padding: 10px 20px;
                        background-color: #ff6f61;
                        color: #ffffff;
                        border: none;
                        border-radius: 5px;
                        cursor: pointer;
                        font-size: 1em;
                    }
                    button[type="submit"]:hover {
                        background-color: #e65c50;
                    }
                    .loading {
                        display: none;
                        margin-top: 20px;
                        font-size: 1.2em;
                        color: #ff6f61;
                    }
                    footer {
                        margin-top: 20px;
                        font-size: 0.8em;
                        color: #999999;
                    }
                    .loader {
                        border: 16px solid #f3f3f3;
                        border-top: 16px solid #ff6f61;
                        border-radius: 50%;
                        width: 120px;
                        height: 120px;
                        animation: spin 2s linear infinite;
                        margin-top: 20px;
                    }
                    table {
                        width: 100%;
                        border-collapse: collapse;
                        margin-top: 20px;
                    }
                    th, td {
                        border: 1px solid #333333;
                        padding: 5px;
                        text-align: center;
                    }
                    th {
                        background-color: #ff6f61;
                        color: #ffffff;
                    }
                    details {
                        margin-top: 10px;
                        text-align: left;
                    }
                    summary {
                        font-weight: bold;
                        cursor: pointer;
                        color: #ff6f61;
                    }
                    summary:hover {
                        text-decoration: underline;
                    }
                    @keyframes spin {
                        0% { transform: rotate(0deg); }
                        100% { transform: rotate(360deg); }
                    }
                </style>
                <script>
                    function showLoading() {
                        document.getElementById('loading').style.display = 'block';
                        document.getElementById('form-input-csv').style.display = 'none';
                        document.querySelector('button[type="submit"]').disabled = true;
                        setTimeout(() => {
                            document.getElementById('loading').style.display = 'none';
                            document.getElementById('form-input-csv').style.display = 'block';
                            document.querySelector('input[type="file"]').value = '';
                            document.querySelector('button[type="submit"]').disabled = false;
                        }, 4000);
                    }
                </script>
            </head>
            <body>
                <div class="container">
                    <div class="row">
                        <h1 class="title">Shop Reporter</h1>
                    </div>
                    <div class="row">
                        <div class="column">
                            <h2>Instrucciones</h2>
                            <details>
                                <summary>1. Estructura del archivo</summary>
                                <p>El archivo debe contener las columnas: tienda, semana, nombre, horas. Como en la <a target="_blank" href="/public/example.csv">plantilla de ejemplo</a>.</p>
                                <table>
                                    <thead>
                                        <tr>
                                            <th>tienda</th>
                                            <th>semana</th>
                                            <th>nombre</th>
                                            <th>horas</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <tr>
                                            <td>1</td>
                                            <td>1</td>
                                            <td>VAZQUEZ SEQUEIRA, ARTURO</td>
                                            <td>40</td>
                                        </tr>
                                        <tr>
                                            <td>1</td>
                                            <td>1</td>
                                            <td>MARTINEZ FULGENCIO, PEPITA</td>
                                            <td>35</td>
                                        </tr>
                                        <tr>
                                            <td>1</td>
                                            <td>2</td>
                                            <td>VAZQUEZ SEQUEIRA, ARTURO</td>
                                            <td>30</td>
                                        </tr>
                                        <tr>
                                            <td>1</td>
                                            <td>2</td>
                                            <td>MARTINEZ FULGENCIO, PEPITA</td>
                                            <td>30</td>
                                        </tr>
                                    </tbody>
                                </table>
                            </details>
                            <details>
                                <summary>2. Formato del archivo</summary>
                                <p>Asegúrate de que el archivo esté en formato <a target="_blank" href="https://support.microsoft.com/es-es/office/importar-o-exportar-archivos-de-texto-txt-o-csv-5250ac4c-663c-47ce-937b-339e391393ba#:~:text=Vaya%20a%20Archivo%20%3E%20Guardar%20como,(delimitado%20por%20comas)..">CSV (delimitado por comas)</a>.</p>
                            </details>
                            <details>
                                <summary>3. Subir el archivo</summary>
                                <p>Haz clic en el botón <i>Subir y procesar</i> para cargar el archivo.</p>
                            </details>
                            <details>
                                <summary>4. Procesamiento</summary>
                                <p>Espera a que el archivo se procese y se genere el reporte.</p>
                            </details>
                            <details>
                                <summary>¿Es seguro el proceso?</summary>
                                <p>¡Claro! Por varios motivos:</p>
                                <ul>
                                    <li>El archivo se procesa en el servidor y no se almacena.</li>
                                    <li>La información se elimina del servidor una vez descargada.</li>
                                    <li>Tanto el proceso de subida como de descarga se realiza a través de HTTPS (cifrado), lo que garantiza la seguridad de los datos.</li>
                                    <li>¡Además! Siempre puedes revisar el código fuente <a href="https://github.com/davgracia/shop-reporter" target="_blank">aquí</a> y descargarlo y ejecutarlo en tu propio ordenador.</li>
                            </details>
                        </div>
                        <div class="column">
                            <h2>Subir Archivo CSV</h2>
                            <p>La transferencia tanto de subida como de descarga de los resultados se hace por canales seguros bajo el protocolo de cifrado HTTPS.</p>
                            <p>La información se procesará y se generará un Libro Excel con los resultados que se descargará automáticamente a tu ordenador.</p>
                            <p><strong>Una vez descargado, toda la información se eliminará del servidor y no será utilizada para otras causas.</strong></p>
                            <form id="form-input-csv" action="/upload" method="post" enctype="multipart/form-data" onsubmit="showLoading()">
                                <input type="file" name="file" accept=".csv" required />
                                <button type="submit">Subir y procesar</button>
                            </form>
                            <p id="loading" class="loading">Subiendo y procesando archivo, por favor espera...</p>
                            <p style="font-size: 0.9em;"><strong>Nota:</strong> Si en tu CSV aparecen datos duplicados como los de la <a target="_blank" href="/public/example.csv">plantilla de ejemplo</a> no te preocupes, el programa elimina las filas duplicadas antes de procesarlas.</p>
                        </div>
                    </div>
                    <footer>
                        &copy; 2025 <a href="https://github.com/davgracia/shop-reporter">SHOP REPORTER</a> (v ${packageJson.version}). Hecho con ❤️ por <a target="_blank" rel="follow" href="https://github.com/davgracia">davgracia</a>.
                    </footer>
                </div>
            </body>
        </html>
    `;
});

router.get('/public/example.csv', async (ctx) => {
    const filePath = path.join(__dirname, 'public', 'example.csv');
    ctx.set('Content-disposition', 'attachment; filename=example.csv');
    ctx.set('Content-type', 'text/csv');
    ctx.body = fs.createReadStream(filePath);
});

router.get('/error', async (ctx) => {
    ctx.body = `
        <html>
            <head>
                <title>Error</title>
                <style>
                    body {
                        background-color: #e74c3c;
                        color: #ecf0f1;
                        text-align: center;
                        font-family: 'Arial', sans-serif;
                        margin: 0;
                        padding: 0;
                    }
                    a {
                        color: #ecf0f1;
                        text-decoration: none;
                        font-weight: bold;
                    }
                    a:hover {
                        text-decoration: underline;
                    }
                    img {
                        width: 300px;
                        margin-top: 20px;
                    }
                    .container {
                        padding: 20px;
                    }
                    h1 {
                        margin-top: 20px;
                        font-size: 2em;
                    }
                    p {
                        font-size: 1.2em;
                        line-height: 1.5;
                    }
                    .button {
                        display: inline-block;
                        padding: 10px 20px;
                        margin-top: 20px;
                        background-color: #c0392b;
                        color: #ecf0f1;
                        border: none;
                        border-radius: 5px;
                        cursor: pointer;
                        font-size: 1em;
                    }
                    .button:hover {
                        background-color: #a93226;
                    }
                </style>
            </head>
            <body>
                <div class="container">
                    <h1>Error al procesar el archivo</h1>
                    <p>Hubo un problema al procesar el archivo. Por favor, inténtalo de nuevo.</p>
                    <a href="/upload" class="button">Volver a subir archivo</a>
                </div>
            </body>
        </html>
    `;
});

router.post('/upload', upload.single('file'), async (ctx) => {
    const file = ctx.file;
    if (!file) {
        ctx.status = 400;
        ctx.body = 'No se ha subido ningún archivo.';
        return;
    }

    const generateRandomString = (length) => {
        const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
        let result = '';
        for (let i = 0; i < length; i++) {
            result += characters.charAt(Math.floor(Math.random() * characters.length));
        }
        return result;
    };

    const randomString = generateRandomString(16) + '.csv';

    const tempPath = file.path;
    const targetPath = path.join(__dirname, './public/input/', randomString);

    fileName = randomString;

    try {
        await fs.promises.rename(tempPath, targetPath);
        ctx.body = `
            <html>
                <head>
                    <title>Shop Reporter</title>
                    <style>
                        body {
                            background-color: green;
                            color: white;
                            text-align: center;
                            font-family: Arial, sans-serif;
                        }
                        a {
                            color: white;
                            text-decoration: underline;
                        }
                    </style>
                </head>
                <body>
                    <h1>Archivo Subido Correctamente</h1>
                    <p>El archivo ${file.originalname} se ha subido correctamente.</p>
                </body>
            </html>
        `;

        const fileProcessed = await processCSV(fileName);

        
        await new Promise(resolve => setTimeout(resolve, 2000));

        if(fileProcessed) {
            ctx.redirect('/error');
        } else {
            ctx.redirect('/download/' + `report_${fileName.replace('.csv', '')}.xlsx`);
        }
    } catch (err) {
        ctx.status = 500;
        ctx.body = 'Error al mover el archivo.';
    }
});

router.get('/download/:fileName', async (ctx) => {
    const fileName = ctx.params.fileName;
    const filePath = path.join(__dirname, './public/output/', fileName);

    if (fs.existsSync(filePath)) {
        if (fs.existsSync(filePath)) {
            ctx.set('Content-disposition', `attachment; filename=${fileName}`);
            ctx.set('Content-type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            ctx.body = fs.createReadStream(filePath);

            ctx.res.on('finish', () => {
                setTimeout(() => {
                    fs.unlink(filePath, (err) => {
                        if (err) {
                            console.error('Error al eliminar el archivo XLSX:', err);
                        } else {
                            console.log('Archivo XLSX eliminado correctamente.');
                        }
                    });
                }, 1000);
            });
        } else {
            ctx.status = 404;
            ctx.body = 'Archivo no encontrado';
        }
    } else {
        ctx.status = 404;
        ctx.body = 'Archivo no encontrado';
    }
});

app.use(router.routes()).use(router.allowedMethods());
app.use(serve(path.join(__dirname, 'public')));

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
    console.log(`Servidor - Shop Reporter escuchando en el puerto ${PORT}`);
});