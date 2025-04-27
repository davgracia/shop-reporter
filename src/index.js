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
            // Elimina los empleados que no tienen horas trabajadas
            const filteredResults = results.filter(entry => parseFloat(entry.horas) > 0);

            const employeeCountByWeek = filteredResults.reduce((acc, entry) => {
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
            const employeeListByWeek = filteredResults.reduce((acc, entry) => {
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
            const weeks = [...new Set(filteredResults.map(entry => entry.semana))];
            const hoursByEmployee = filteredResults.reduce((acc, entry) => {
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
            const employeeRanking = filteredResults.reduce((acc, entry) => {
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
                employee.semanas = filteredResults.reduce((acc, entry) => {
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
                const weeklyHours = filteredResults
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
                const weeklyHours = filteredResults
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
                    @import url('https://fonts.googleapis.com/css2?family=Nunito:wght@400;600&display=swap');

                    body {
                        background-color: #f5f5f5;
                        color: #2c3e50;
                        font-family: 'Nunito', sans-serif;
                        margin: 0;
                        padding: 0;
                        display: flex;
                        flex-direction: column;
                        align-items: center;
                        justify-content: center;
                        min-height: 100vh;
                        text-align: justify; /* Justify text */
                    }

                    a {
                        color: #2c3e50;
                        text-decoration: none;
                        font-weight: 600;
                        transition: color 0.3s;
                    }

                    a:hover {
                        color: #1a252f;
                    }

                    img {
                        width: 100px;
                        margin: 20px 0;
                    }

                    .container {
                        width: 90%;
                        max-width: 1000px;
                        background: #ffffff;
                        padding: 40px;
                        border-radius: 12px;
                        box-shadow: 0 4px 20px rgba(0,0,0,0.05);
                    }

                    .row {
                        display: flex;
                        flex-wrap: wrap;
                        gap: 20px;
                        justify-content: center;
                    }

                    .title {
                        font-size: 2.2em;
                        font-weight: 600;
                        text-align: center;
                        margin-bottom: 30px;
                        color: #2c3e50;
                        border-bottom: 2px solid #e8e4e1;
                        padding-bottom: 10px;
                    }

                    .column {
                        flex: 1 1 300px;
                        background-color: #f5f5f5;
                        border-radius: 8px;
                        padding: 20px;
                        text-align: center;
                        transition: transform 0.3s;
                    }

                    .column:hover {
                        transform: translateY(-5px);
                    }

                    h1 {
                        font-size: 1.5em;
                        margin-bottom: 10px;
                        color: #2c3e50;
                    }

                    p {
                        font-size: 1em;
                        line-height: 1.6;
                        color: #666666;
                    }

                    .button {
                        background-color: #ffffff;
                        color: #2c3e50;
                        padding: 12px 24px;
                        border: 2px solid #2c3e50;
                        border-radius: 50px;
                        font-size: 1em;
                        margin-top: 20px;
                        cursor: pointer;
                        transition: background 0.3s, color 0.3s, border-color 0.3s;
                    }

                    .button:hover {
                        background-color: #2c3e50;
                        color: #ffffff;
                        border-color: #ffffff;
                    }

                    input[type="file"] {
                        margin-top: 20px;
                        border: 2px dashed #2c3e50;
                        padding: 20px;
                        background: #ffffff;
                        border-radius: 8px;
                        width: 100%;
                        text-align: center;
                        font-size: 1em;
                        cursor: pointer;
                        transition: background-color 0.3s;
                    }

                    input[type="file"]:hover {
                        background-color: #f0f0f0;
                    }

                    button[type="submit"] {
                        background-color: #ffffff;
                        color: #2c3e50;
                        padding: 12px 24px;
                        border: 2px solid #2c3e50;
                        border-radius: 50px;
                        margin-top: 20px;
                        font-size: 1em;
                        cursor: pointer;
                        transition: background 0.3s, color 0.3s, border-color 0.3s;
                        width: 100%; /* Full width */
                    }

                    button[type="submit"]:hover {
                        background-color: #2c3e50;
                        color: #ffffff;
                        border-color: #ffffff;
                    }

                    button[type="submit"]:disabled {
                        background-color: #cccccc;
                        color: #666666;
                        border-color: #cccccc;
                        cursor: not-allowed;
                    }

                    .loading {
                        margin-top: 20px;
                        font-size: 1.2em;
                        color: #2c3e50;
                        display: none;
                    }

                    .loader {
                        border: 6px solid #f0f0f0;
                        border-top: 6px solid #2c3e50;
                        border-radius: 50%;
                        width: 60px;
                        height: 60px;
                        animation: spin 1s linear infinite;
                        margin: 20px auto;
                    }

                    table {
                        width: 100%;
                        border-collapse: collapse;
                        margin-top: 20px;
                        font-size: 0.9em;
                    }

                    th, td {
                        border: 1px solid #e8e4e1;
                        padding: 10px;
                    }

                    th {
                        background-color: #2c3e50;
                        color: #ffffff;
                    }

                    details {
                        margin-top: 20px;
                        background: #f5f5f5;
                        padding: 15px;
                        border-radius: 8px;
                    }

                    summary {
                        font-weight: 600;
                        color: #2c3e50;
                        cursor: pointer;
                        outline: none;
                    }

                    footer {
                        margin-top: 40px;
                        font-size: 0.8em;
                        color: #999999;
                        text-align: center;
                    }

                    @keyframes spin {
                        0% { transform: rotate(0deg); }
                        100% { transform: rotate(360deg); }
                    }

                    .modal {
                        display: none;
                        position: fixed;
                        z-index: 1000;
                        left: 0;
                        top: 0;
                        width: 100%;
                        height: 100%;
                        overflow: auto;
                        background-color: rgba(0, 0, 0, 0.5);
                        display: flex;
                        align-items: center;
                        justify-content: center;
                    }

                    .modal-content {
                        background-color: #fff;
                        padding: 20px;
                        border: 1px solid #888;
                        width: 90%;
                        max-width: 600px;
                        border-radius: 8px;
                        text-align: center;
                    }

                    .close {
                        color: #aaa;
                        float: right;
                        font-size: 28px;
                        font-weight: bold;
                        cursor: pointer;
                    }

                    .close:hover,
                    .close:focus {
                        color: black;
                        text-decoration: none;
                    }
                </style>

                <script>
                    function openModal() {
                        document.getElementById('infoModal').style.display = 'flex';
                    }

                    function closeModal() {
                        document.getElementById('infoModal').style.display = 'none';
                    }

                    window.onclick = function(event) {
                        const modal = document.getElementById('infoModal');
                        if (event.target === modal) {
                            modal.style.display = 'none';
                        }
                    };

                    function toggleSubmitButton() {
                        const fileInput = document.querySelector('input[type="file"]');
                        const submitButton = document.querySelector('button[type="submit"]');
                        submitButton.disabled = !fileInput.files.length;
                    }
                </script>
            </head>
            <body>
                <div class="container">
                    <div class="row">
                        <h1 class="title">🗓️ Shop Reporter</h1>
                    </div>
                    <div class="row">
                        <div id="infoModal" class="modal" style="display: none;">
                        <div class="modal-content">
                            <span class="close" onclick="closeModal()">&times;</span>
                            <h2>Información</h2>
                            <p>Esta herramienta está diseñada para procesar archivos CSV de manera segura. Los archivos subidos son procesados y eliminados automáticamente después de generar el reporte, minimizando el riesgo de almacenamiento innecesario. Solo se aceptan archivos con extensión <code>.csv</code> para evitar la ejecución de archivos maliciosos.</p>
                            <p>Si tienes dudas o inquietudes sobre la seguridad, por favor, recuerda que el <a href="https://github.com/davgracia/shop-reporter" target="_blank">código fuente de esta aplicación</a> es público. ¡Puedes examinarlo o descargarlo en tu ordenador y ejecutarlo ahí!
                        </div>
                    </div>
                    <div class="column">
                        <h1>📄 ¿Cómo funciona?</h1>
                        <p>Sube un archivo CSV con los datos de las tiendas y empleados como en el <a href="/public/example.csv">archivo de ejemplo</a>. El sistema procesará la información y generará un reporte detallado en formato Excel que descargará en tu ordenador de forma automática.</p>

                        <button class="button" onclick="openModal()">ℹ️ Más info</button>
                        <a href="/public/example.csv" class="button">📥 Archivo de ejemplo</a>
                    </div>
                    <div class="column">
                        <h1>🚀 ¡Vamos!</h1>
                        <form id="form-input-csv" action="/upload" method="post" enctype="multipart/form-data" onsubmit="showLoading()">
                            <input type="file" name="file" accept=".csv" required onchange="toggleSubmitButton()" />
                            <button type="submit" disabled>Procesar</button>
                        </form>
                    </div>
                    <footer>
                        &copy; 2025 <a href="https://github.com/davgracia/shop-reporter">SHOP REPORTER</a> (v ${packageJson.version}). Hecho con 💙 por <a target="_blank" rel="follow" href="https://github.com/davgracia">davgracia</a>.
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