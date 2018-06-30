$(document).ready(function () {

    //var url_string = window.location.href;
    //var url = new URL(url_string);
    //var c = url.searchParams.get("c");

    var c = getParameterByName("c");

    $.ajax({
        url: 'handlers/ArchivosCotizacion.ashx',
        data: { 1: 'dataQuotation', quotationId: c },
        success: function (result) {
            if (result.done) {


                $('#titulo').text('ARCHIVOS DE LA COTIZACION Nº ' + result.data.IT);

            }
        }
    })


    //updateQuotation(c);
    cargarFiles(c);

});

function updateQuotation(quotationId) {
    $.ajax({
        url: 'handlers/ArchivosCotizacion.ashx',
        data: { 1: 'updateQuotation', quotationId: quotationId },
        success: function (result) {
            if (!result.done) {
                console.log(result);
                return;
            }
        }
    })

}

function cargarFiles(quotationId) {
    $.ajax({
        url: 'handlers/ArchivosCotizacion.ashx',
        data: { 1: 'getFiles', quotationId: quotationId },
        success: function (result) {
            if (!result.done) {
                console.log(result);
                return;
            }
            if (result.data.length == 0) {
                $('#tituloPanel').text('No hay Archivos para la cotizacion Consultada.');
                return;
            }
            $(result.data).each(function (i, item) {
                $('#filesPanel')
                    .append($('<div>')
                        .data('id', item.id)
                        .click(function () {

                            var folder = item.isQuotation ? 'cotizaciones' : 'archivos_mail'
                            //download('/certel/' + folder + '/' + item.url, item.url)
                            download('/' + folder + '/' + item.url, item.url)
                            //if(navigator.userAgent.toLowerCase().indexOf('firefox') > -1)
                              //  window.location.assign(folder + '/' + item.url)
                            //else
                              //  download('/' + folder + '/' + item.url, item.url)
                            if (item.isQuotation) {
                                updateQuotation(quotationId);
                            }
                        })
                        .addClass(item.isQuotation ? 'fileSpecial' : 'file')
                        //.css(item.isQuotation ? { background: "#88CA7B"} : {})
                        .addClass('col-xs-12 col-lg-3 col-md-4 col-sm-6')
                        .text(item.nombre)
                        .append($('<i>')
                            .addClass(item.extencion == 'pdf' ? 'fa fa-file-pdf-o' : (item.extencion == 'docx' || item.extencion == 'doc') ? 'fa fa-file-word-o' : (item.extencion == 'xlsx' || item.extencion == 'xls') ? 'fa fa-file-excel-o' : 'fa fa-file-text-o')))



            });
        }
    })

}

function download(uri, name) {
    if(navigator.userAgent.toLowerCase().indexOf('firefox') > -1) {
        window.open(uri, 'Download')
        return;
    }
    var link = document.createElement("a")
    link.download = name;
    link.href = uri;
    document.body.appendChild(link)
    link.click();
    document.body.removeChild(link)
}

