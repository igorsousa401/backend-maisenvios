const xlsx = require('xlsx');
const express = require('express');
const _ = require('lodash');
const app = express();
const port = 2000;

//Rota que retona todos os registros que estão na planilha
app.get('/etiquetas', (req, res) => {
    const wb = xlsx.readFile(file);
    const ws = wb.Sheets["Plan1"];
    const rows = xlsx.utils.sheet_to_json(ws);

    let data = [];

    for(let index = 4; index < rows.length + 3; index++) {
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

//Rota Delete, para deletar é necessário você passar o número da row ao qual você quer deletar como parâmetro. Ex.: '/etiquetas/4'. Você só pode deletar rows que envolvam dados, ou seja, somente as rows que envolvam as etiquetas de fato."
app.delete('/etiquetas/:index', (req, res) => {
    const wb = xlsx.readFile(file);
    const ws = wb.Sheets["Plan1"];
    const rows = xlsx.utils.sheet_to_json(ws);
    req.params.index = req.params.index - 1;
    function ec(r, c) {
        return XLSX.utils.encode_cell({ r: r, c: c });
    }

    if(3 < req.params.index <= rows.length + 3) {
        var decode = xlsx.utils.decode_range(ws["!ref"])
        for (var reqI = req.params.index; reqI < decode.e.r; ++reqI) {
            for (var C = decode.s.c; C <= decode.e.c; ++C) {
                console.log(C);
                ws[xlsx.utils.encode_cell({ r: reqI, c: C })] = ws[xlsx.utils.encode_cell({ r: reqI+1, c: C })];
            }
        }
        decode.e.r--;
        ws['!ref'] = xlsx.utils.encode_range(decode.s, decode.e);
        xlsx.writeFile(wb, file);
        res.send({
            'success' : true,
            'message': "A Row foi deletada com sucesso"
        });
    } else {
        res.send({
            'success' : false,
            'message': "A row passada não é um valor aceito. Confira novamente a planilha e verifique se está tentando deletar a correta."
        });
    }



});

//Rota Delete, para deletar é necessário você passar o número da row ao qual você quer deletar como parâmetro. Ex.: '/etiquetas/4'. Você só pode deletar rows que envolvam dados, ou seja, somente as rows que envolvam as etiquetas de fato."
app.put('/etiquetas/:index', (req, res) => {
    const wb = xlsx.readFile(file);
    const ws = wb.Sheets["Plan1"];
    const rows = xlsx.utils.sheet_to_json(ws);

        res.send({
            'success' : false,
            'message': "A row passada não é um valor aceito. Confira novamente a planilha e verifique se está tentando deletar a correta."
        });




});

app.listen(port, () => {
    console.log(`O servidor está aberto na porta ${port}`);
})

const file = './lista_etiquetas.xlsx';
