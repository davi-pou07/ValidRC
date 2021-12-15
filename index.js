const express = require("express")
const app = express()
const path = require('path')
const bodyParser = require("body-parser")
const XLSX = require('node-xlsx')
const ejs = require("ejs")
const fs = require('fs')
const multer = require("multer");
const validator = require("validator")
const excel = require('node-excel-export');

const xl = require('excel4node');
const wb = new xl.Workbook();
const ws = wb.addWorksheet('Worksheet Name');

const validaTelefone = require("./public/js/validaTelefone.js")
const validaEmail = require("./public/js/validaEmail.js")

app.set('views', path.join(__dirname, 'views'))
app.set('view engine', 'ejs')

const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, 'public/upload/')
    },
    filename: function (req, file, cb) {
        const filename = Date.now().toString()
        const extensionFile = file.originalname.slice(file.originalname.lastIndexOf("."));
        cb(null, `${filename}${extensionFile}`);
    }
})

const upload = multer({ storage })

//Carregamento de arquivos estaticos no express
app.use(express.static(path.join(__dirname, 'public')))
//Carregamento do bodyPerser
app.use(bodyParser.urlencoded({ extended: false, limit: "50mb" }))
app.use(bodyParser.json({ limit: '50mb' }))

app.get("/", (req, res) => {
    var dados = 0
    res.render("index", { dados: dados })

})

app.post("/exportar", upload.single("arquivo"), (req, res) => {
    var filename = req.file.filename
    res.redirect(`/${filename}`)
})

app.get("/:arquivo", (req, res) => {
    var arquivo = req.params.arquivo
        var dados = []
        if (arquivo != undefined) {
            var obj = XLSX.parse(fs.readFileSync(`./public/upload/${arquivo}`));
            var tabela = obj[0].data
            tabela.forEach(linha => {
                if (linha[4] != undefined || linha[5] != undefined) {
                    var phone1 = validaTelefone(linha[4])
                    var phone2 = validaTelefone(linha[5])
                }
                var email = validaEmail(linha[3])
                dados.push({ id: linha[0], nome: linha[1], cgc: linha[2], email: email, tel1: phone1, tel2: phone2 })
            })
        } else {
            dados = 0
        }
        res.json({ dados: dados })
})

app.get("/teste/:arquivo", (req, res) => {
    var arquivo = req.params.arquivo
    var dados = []
    if (arquivo != undefined) {
        var obj = XLSX.parse(fs.readFileSync(`./public/upload/${arquivo}`));
        var tabela = obj[0].data
        tabela.forEach((linha, index) => {
            if (index != 0 && index<1000) {
                function validaTelefone2(ddd, numero) {
                    var phone = []
                    if (numero != undefined && ddd != undefined) {
                        numero = numero.toString()
                        ddd = ddd.toString()
                        var nm = numero.split("")
                        var dd = ddd.split("")
                        if (nm[0] == dd[0] && nm[1] == dd[1]) {
                            numero = numero.split(ddd)[1]
                        }
                        if (numero == undefined) {
                            numero = 'NULL'
                        }
                        var vl = numero.split("")
                        if (vl[0] == 6 || vl[0] == 7 || vl[0] == 8 || vl[0] == 9) {
                            if (numero.length == 8) {
                                var n = `${ddd}9${numero}`
                            } else {
                                var n = `${ddd}${numero}`
                            }
                            var status = validator.isMobilePhone(n, 'pt-BR', false)
                            var is = 'CEL'
                        } else if (vl[0] == 2 || vl[0] == 3 || vl[0] == 4 || vl[0] == 5) {
                            var n = `${ddd}${numero}`
                            var is = 'FIX'
                            var status = validator.isMobilePhone(n, 'pt-BR', false)
                        }
                    } else {
                        var n = `${ddd}${numero}`
                    }

                    phone.push({ numero: n, status: status, is: is })
                    return phone
                }
                if (linha[4] == undefined) {
                    linha[4] = 'NULL'
                }
                if (linha[5] == undefined) {
                    linha[5] = 'NULL'
                }
                if (linha[6] == undefined) {
                    linha[6] = 'NULL'
                }
                if (linha[7] == undefined) {
                    linha[7] = 'NULL'
                }
                var phone1 = validaTelefone2(linha[4], linha[5])

                var phone2 = validaTelefone2(linha[6], linha[7])
                var email = validaEmail(linha[3])

                dados.push({ id: linha[0], nome: linha[1], cgc: linha[2], email: email, tel1: phone1, tel2: phone2 })
            }
        })
    } else {
        dados = 0
    }
    res.render('index', { dados: dados })
})
//TESTE

app.listen(8080, () => {
    console.log("Servidor ligado")
})