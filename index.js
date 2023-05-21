const xlsx = require('xlsx');
const express = require('express');
const _ = require('lodash');
const app = express();
const port = 2000;
const file = './lista_etiquetas.xlsx';

/*
*** Rota GET
* Para mostrar os dados da planilha é bem simples, é só enviar uma requisição GET para esta Rota:"http://localhost:2000/etiquetas/".
* Lembre-se de selecionar o method DELETE.
*/
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

/*
*** Rota DELETE
* Para Deletar a etiqueta é nessário passar o número da sua row, Ex.: "http://localhost:2000/etiquetas/2".
* Lembre-se de selecionar o method DELETE.
*/
app.delete('/etiquetas/:index', (req, res) => {
    const wb = xlsx.readFile(file);
    const ws = wb.Sheets["Plan1"];
    const rows = xlsx.utils.sheet_to_json(ws);
    req.params.index = req.params.index - 1;
    function ec(r, c) {
        return xlsx.utils.encode_cell({ r: r, c: c });
    }

    if(3 < req.params.index && req.params.index <= rows.length + 1) {
        var decode = xlsx.utils.decode_range(ws["!ref"])
        for (var reqI = req.params.index; reqI < decode.e.r; ++reqI) {
            for (var C = decode.s.c; C <= decode.e.c; ++C) {
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

/*
*** Rota PUT
* Para Atualizar o valor da etiqueta é nessário passar o número da sua row, desta forma, Ex.:"http://localhost:2000/etiquetas/2".
* Para atualizar a row é necessário passar como parâmetros query os valores que deseja atualizar, sendo esses: "tag", "name", "status", "price" e "source".
* A passagem dos parâmetros é opcional, então escolha somente aqueles que deseja alterar, lembrando de passar na url o número correto da row que deseja alterar.
* * Lembre-se de selecionar o method PUT.
*/
app.put('/etiquetas/:index', (req, res) => {
    const wb = xlsx.readFile(file);
    const ws = wb.Sheets["Plan1"];
    const rows = xlsx.utils.sheet_to_json(ws);

    if(3 < req.params.index && req.params.index <= rows.length + 1) {
        for(var reqI = 4; reqI < rows.length + 3; reqI++) {
            if(req.params.index == reqI) {
                const tag = req.query.tag != null ? req.query.tag : ws[`A${reqI}`].v;
                ws[`A${reqI}`].v = tag;
                const name = req.query.name != null ? req.query.name : ws[`B${reqI}`].v;
                ws[`B${reqI}`].v = name;
                const status = req.query.status != null ? req.query.status : ws[`C${reqI}`].v;
                ws[`C${reqI}`].v = status;
                const source = req.query.source != null ? req.query.source : ws[`D${reqI}`].v;
                ws[`D${reqI}`].v = source;
                const price = req.query.price != null ? req.query.price : ws[`E${reqI}`].v;
                ws[`E${reqI}`].v = price;

                xlsx.writeFile(wb, file);

                res.send({
                    'success' : true,
                    'message': "A Row foi atualizada com sucesso"
                });
            }

        }

    } else {
        res.send({
            'success' : false,
            'message': "A row passada não é um valor aceito. Confira novamente a planilha e verifique se está tentando atualizar a row correta."
        });
    }
});

app.listen(port, () => {
    console.log(`O servidor está aberto na porta ${port}`);
})