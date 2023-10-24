
const express = require('express');
const fs = require('fs');
const path = require('path');
const exceljs = require('exceljs');

const router = express.Router();

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
            sheet.addRow({ id: value.id, endereco: value.endereco, status: value.status });
        });

        const buffer = await workbook.xlsx.writeBuffer();
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=casos.xlsx');
        res.send(buffer);
    } catch (error) {
        console.error(error);
        res.status(500).send('Error exporting data to Excel');
    }
});

module.exports = router;