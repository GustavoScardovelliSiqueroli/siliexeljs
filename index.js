const express = require('express');
const exceljs = require('exceljs');
var fs = require('fs');

const app = express();

const PORT = 3000;

const path = require('path');
const { title } = require('process');
const router = express.Router();

router.get('/', (req, res) => {

    res.sendFile(path.join(__dirname + '/index.html'));
});

router.get('/export', async (req, res) => {
    try {
        let workbook = new exceljs.Workbook();

        const sheet = workbook.addWorksheet("casos")
        sheet.columns = [
            { header: "ID", key: "id", width: 25 },
            { header: "Endereço", key: "endereco", width: 25 },
            { header: "Status", key: "status", width: 25 },
        ];

        let object = JSON.parse(fs.readFileSync('cases.json', 'utf8'));

        await object.map((value, idx) => {
            sheet.addRow({ id: value.id,
                 endereco: value.endereco,
                  status: value.status,
                 });
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader("Content-Disposition", "attachment; filename=" + "cases.xlsx");
        workbook.xlsx.write(res)

    } catch (error) {
        console.log(error);
    }

})

app.use('/', router);

app.listen(PORT, () => {
    console.log(`Server is runing on port: ${PORT}`);
});

