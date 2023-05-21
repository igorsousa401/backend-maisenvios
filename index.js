const xlsx = require('xlsx');
const express = require('express');
const _ = require('lodash');
const app = express();
const port = 2000;

app.get('/', (req, res) => {
    const wb = xlsx.readFile(file);
    const ws = wb.Sheets["Plan1"];
    const rows = xlsx.utils.sheet_to_json(ws);

    const spec = {};
    const _u = _.noConflict();

    let data = [];

    for(let index = 4; index < 10; index++) {
        const tag = ws[`A${index}`].v;
        const name = ws[`B${index}`].v;
        const status = ws[`C${index}`].v;
        const source = ws[`D${index}`].v;
        const price = ws[`E${index}`].v;
        let row = {
            tag:tag,
            name:name,
            status:status,
            source:source,
            price:price
        };
        data.push(row);
    }
    data = JSON.parse(JSON.stringify(data))
    res.send(data);

});

app.listen(port, () => {
    console.log(`O servidor está aberto na porta ${port}`);
})

const file = './lista_etiquetas.xlsx';
