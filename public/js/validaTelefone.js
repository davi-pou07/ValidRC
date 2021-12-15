const validator = require("validator")


function validaTelefone(numero) {
    var phone = []
    if (numero != 'tel1' && numero != 'tel2' && numero != undefined) {
        numero = numero.toString()
        function dd() {
            var sp = numero.split("")
            if (sp[0] == 0) {
                var dd = numero.split("")[1] + numero.split("")[2]
            } else {
                var dd = numero.split("")[0] + numero.split("")[1]
            }
            return dd
        }
        var ddd = dd()
        var num = numero.split(ddd)[1]
        if (num != undefined) {
            var vl = num.split("")
            if (vl[0] == 6 || vl[0] == 7 || vl[0] == 8 || vl[0] == 9) {
                if (num.length == 8) {
                    var n = `${ddd}9${num}`
                } else {
                    var n = `${ddd}${num}`
                }
                var status = validator.isMobilePhone(n, 'pt-BR', false)
                var is = 'CEL'
            } else if (vl[0] == 2 || vl[0] == 3 || vl[0] == 4 || vl[0] == 5) {
                var n = `${ddd}${num}`
                var is = 'FIX'
                var status = validator.isMobilePhone(n, 'pt-BR', false)
            }
        } else {
            n = numero
            status = false
            is = 'NULL'
        }
    } else {
        n = numero
        status = false
        is = 'NULL'
    }
    phone.push({ numero: n, status: status, is: is })
    return phone

}

module.exports = validaTelefone
