const validator = require("validator")


function validaEmail(ema) {
    var em = []
    var status = ''
    if (ema != undefined) {
        var isemail = validator.isEmail(ema)
        if (isemail == true) {
            var email = ema.split('@')
            if (email[0].includes('&')) {
                status = 'invalido'
            }

            if (email[0].includes('=')) {
                status = 'invalido'
            }

            if (email[0].includes("'")) {
                status = 'invalido'
            }

            if (email[0].includes("-")) {
                status = 'invalido'
            }

            if (email[0].includes("+")) {
                status = 'invalido'
            }

            if (email[0].includes(",")) {
                status = 'invalido'
            }

            if (email[0].includes("<")) {
                status = 'invalido'
            }

            if (email[0].includes(">")) {
                status = 'invalido'
            }

            if (email[0].includes("..")) {
                status = 'invalido'
            }

            if (email[0].includes("/")) {
                status = 'invalido'
            }
            if (email[1].includes("grupo")) {
                status = 'invalido'
            }
        } else {
            status = 'invalido'
        }
        if (status != 'invalido') {
            status = 'valido'
        }
    } else {
        status = 'invalido'
    }
    em.push({ email: ema, status: status })
    return em
}

module.exports = validaEmail