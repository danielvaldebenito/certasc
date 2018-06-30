$(document).ready(function () {
    $('#user').focus();
    $('#form-login').submit(function (e) {
        e.preventDefault();
        $.ajax({
            url: 'handlers/Login.ashx',
            method: 'post',
            data: {
                user: $('#user').val(),
                pass: $('#pass').val(),
            },
            success: function (result) {
                console.log(result);
                if(result.done)
                {
                    getValuesToday();
                    localStorage.setItem('user', result.user);
                    localStorage.setItem('name', result.name);
                    localStorage.setItem('roles', JSON.stringify(result.roles));
                    setTimeout(function () {
                        console.log('redirecting...')
                        window.location.href = 'Inspecciones.aspx';
                    }, 1000)
                    
                }
                else
                {
                    alertify.error(result.message);
                    $('#pass').val('').focus();
                }
            }

        })
    });

});

function getValuesToday()
{
    $.ajax({
        url: 'http://api.sbif.cl/api-sbifv3/recursos_api/uf?apikey=10cef36a2707ccaf2cf71ec5b8884fd418ee9989&formato=json',
        success: function (result) {
            var uf = result.UFs;
            if (uf && uf[0]) {
                localStorage.setItem('UF', uf[0].Valor)
                localStorage.setItem('DateUF', uf[0].Fecha)
                saveMoneda('UF', uf[0].Valor)
            }
        }
    })
    $.ajax({
        url: 'http://api.sbif.cl/api-sbifv3/recursos_api/dolar?apikey=10cef36a2707ccaf2cf71ec5b8884fd418ee9989&formato=json',
        success: function (result) {
            var dolar = result.Dolares;
            if (dolar && dolar[0]) {
                localStorage.setItem('Dolar', dolar[0].Valor)
                localStorage.setItem('DateDolar', dolar[0].Fecha)
                saveMoneda('Dolar', dolar[0].Valor)
            }
        }
    })
    $.ajax({
        url: 'http://api.sbif.cl/api-sbifv3/recursos_api/euro?apikey=10cef36a2707ccaf2cf71ec5b8884fd418ee9989&formato=json',
        success: function (result) {
            var euro = result.Euros;
            if (euro && euro[0]) {
                localStorage.setItem('Euro', euro[0].Valor)
                localStorage.setItem('DateEuro', euro[0].Fecha)
                saveMoneda('Euro', euro[0].Valor)
            }
        }
    })
}

function saveMoneda(type, value) {
    $.ajax({
        url: 'handlers/Config.ashx',
        data: { 1: 'setMoneda', type: type, value: value },
        success: function (result) {
            console.log(result)
        }
    })
}