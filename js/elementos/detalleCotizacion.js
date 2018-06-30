var GLOBAL;
var MailCliente;
var contenidoMAIL;
var COTIZACION;
$(document).ready(function () {

    //var url_string = window.location.href;
    //var url = new URL(url_string);
    //var c = url.searchParams.get("c");

    var c = getParameterByName("c");
    $('#correo1').jqte()
    createInitialView(c);
    getGlobal();
    $('#add-fechaCompromiso').datetimepicker({
        locale: 'es',
        format: 'DD-MM-YYYY HH:mm'
    });

    // TOOLTIPS
    $('#send-mail, #add-hito, #reenviar').tooltip();
    combobox($('#add-tipo-hito'), { 1: 'hitos', quotationId: c }, 'Seleccione...');
    combobox($('#add-ingeniero'), { 1: 'ingenieros' }, 'Seleccione...');

    $('#correo').jqte();


    $('#divIngeniero').hide('fast');
    $('#divFechaCompromiso').hide('fast');

    $('#add-tipo-hito').change(function () {
        
        var showDateArray = ['13', '17','19', '21'];
        console.log(showDateArray.indexOf($(this).val()))
        var showDate = showDateArray.indexOf($(this).val()) > -1;
        
        if (showDate) {
            $('#divIngeniero').show('fast');
            $('#divFechaCompromiso').show('fast');
            $("#add-ingeniero").attr("required", true);
            $("#add-fechaCompromiso").attr("required", true);
            if ($(this).val() == '21' || $(this).val() == '19') {
                $("#add-ingeniero").val(COTIZACION.ingeniero).prop('disabled', true);
            } else {
                $("#add-ingeniero").val(null).prop('disabled', false);
            } 

        } else {
            $('#divIngeniero').hide('fast');
            $('#divFechaCompromiso').hide('fast');
            $("#add-ingeniero").attr("required", false);
            $("#add-fechaCompromiso").attr("required", false);
        }

    });


    $("#grid").jqGrid({
        url: 'handlers/DetalleCotizacion.ashx',
        postData: {
            1: 'grid',
            id: c
        },
        mtype: "POST",
        autowidth: true,
        datatype: "json",
        shrinkToFix: true,
        ajaxGridOptions: { global: false },
        colModel: [
            { label: 'Id', name: 'Id', key: true, hidden: true },
            { label: 'Fecha Hito', name: 'Fecha', index: 'Fecha', width: 15, align: 'left', sortable: false },
            { label: 'Hito', name: 'Hito', index: 'Hito', width: 30, align: 'left', sortable: true },
            { label: 'Observacion', name: 'Observacion', index: 'Observacion', width: 35, align: 'left', sortable: true },
            { label: 'Fecha Compromiso', name: 'FechaCompromiso', index: 'FechaCompromiso', width: 30, align: 'left', sortable: true },
        ],
        sortname: 'FechaCreacion', sortorder: 'desc',
        viewrecords: true,
        autoheight: true,
        rowNum: 30,
        pager: "#pager",
        caption: '',

        onCellSelect: function (rowid, iCol, cellcontent, e) {
            var rowData = $(this).jqGrid('getRowData', rowid);
            SELECTED = rowData;
        }
    }).navGrid('#pager',
        {
            edit: false, add: false, del: false, search: false,
            refresh: true, view: false, position: "left", cloneToTop: false
        });

    $('#add-hito').click(function () {

        $('#add-hito-dialog').dialog({
            modal: true, bgiframe: false, width: 600, title: 'Nuevo Hito', draggable: true,
            resizable: false, closeOnEscape: true, autoOpen: true, show: 'clip', hide: 'puff',
            open: function (type, data) {
                $(this).find('form')[0].reset();
            },
            close: function () {

            },
            buttons: [
                {
                    id: 'btn-add',
                    text: 'Guardar',
                    click: function () {
                        $('#add-hito-form').submit();
                    }
                }
            ]
        });
    });

    $('#add-hito-form').submit(function (e) {
        e.preventDefault();
        if (!validateForm($(this))) {
            alertify.error('Complete los datos requeridos');
            return false;
        }

        $.ajax({
            url: 'handlers/DetalleCotizacion.ashx',
            data: {
                1: 'addHito',
                quotationId: c,
                tipo: $('#add-tipo-hito').val(),
                observacion: $('#add-obs').val(),
                fechaCompromiso: $('#add-fechaCompromiso').val(),
                ingeniero: $('#add-ingeniero').val(),
                user: localStorage.getItem('user')
            },
            success: function (result) {
                if (result.done) {
                    alertify.success(result.message);
                    //$('#add-hito-dialog').dialog('close');

                    //reloadGridHitos(c);
                    //createInitialView(c);

                    //$('#divIngeniero').hide('fast');
                    //$('#divFechaCompromiso').hide('fast');
                    //$("#add-ingeniero").attr("required", false);
                    //$("#add-fechaCompromiso").attr("required", false);

                    //combobox($('#add-tipo-hito'), { 1: 'hitos', quotationId: c }, 'Seleccione...');
                    //combobox($('#add-ingeniero'), { 1: 'ingenieros' }, 'Seleccione...');
                    location.reload();
                }
                else
                    alertify.error(result.message);
            }
        });
        return false;

    });

    $('#send-mail').click(function () {
        if ($('#correo').val() != '')
            contenidoMAIL = $('#correo').val()

        $('#send-mail-dialog').dialog({
            modal: true, bgiframe: false, width: 700, title: 'Nuevo Correo', draggable: true,
            resizable: false, closeOnEscape: true, autoOpen: true, show: 'clip', hide: 'puff',
            open: function (type, data) {
                $(this).find('form')[0].reset();
                $('#mail-destinatario').val(MailCliente);
            },
            close: function () {

            },
            buttons: [
                {
                    id: 'btn-send',
                    text: 'Enviar',
                    click: function () {
                        $('#send-mail-form').submit();
                    }
                }
            ]
        });
    });

    $('#send-mail-form').submit(function (e) {
        e.preventDefault();
        if (!validateForm($(this))) {
            alertify.error('Complete los datos requeridos');
            return false;
        }
        console.log("textarea", $('#correo').val());
        console.log("Variable", contenidoMAIL);
        $.ajax({
            url: 'handlers/DetalleCotizacion.ashx',
            data: {
                1: 'sendEmail',
                quotationId: c,
                destinatarios: $('#mail-destinatario').val(),
                contenido: $('#correo').val() == '' ? contenidoMAIL : $('#correo').val(),
                asunto: $('#mail-asunto').val(),
                user: localStorage.getItem('user')
            },
            success: function (result) {
                if (result.done) {
                    alertify.success(result.message);
                    window.location.reload();
                }
                else
                    alertify.error(result.message);
            }
        });
        return false;

    });

    $('#ShowCalendar').click(function () {
        window.open('calendar.aspx?user=' + $('#add-ingeniero').val() + "&name=" + $('#add-ingeniero option:selected').text(), '_blank', 'fullscreen=yes')
    });

    createItemsView(c);
    createCostosLogisticosView(c);

    // DIALOGS
    $('#reenviar').click(function () {
        $('#reenviar-dialog').dialog({
            modal: true, bgiframe: false, width: 700, title: 'Reenviar Cotización', draggable: true,
            resizable: false, closeOnEscape: true, autoOpen: true, show: 'clip', hide: 'puff',
            open: function (type, data) {
                $(this).find('form')[0].reset();
                $('#mail-destinatario').val(MailCliente);
                $('#para').val(MailCliente);
                $('#asunto').val(GLOBAL.mail_envio_cotizacion_subject);
                $('#correo1').jqteVal(GLOBAL.mail_envio_cotizacion_text);
            },
            close: function () {

            },
            buttons: [
                {
                    id: 'btn-send',
                    text: 'Enviar',
                    click: function () {
                        $('#reenviar-form').submit();
                    }
                }
            ]
        });
    })
    $('#add-para').click(function () {
        addPara();
    })
    $('#add-cc').click(function () {
        addCC();
    })
    $('#add-cco').click(function () {
        addCCO();
    })
    $('#reenviar-form').submit(function () {
        $('#reenviar-dialog').dialog('close');
        $.ajax({
            url: 'handlers/Cotizaciones.ashx',
            type: 'POST',
            global: false,
            data: {
                1: 'reenviar',
                cotizacionId: c,
                asunto: $('#asunto').val(),
                para: $('.para').serializeArray().map(function (m) { return m.value }),
                cc: $('.cc').serializeArray().map(function (m) { return m.value }),
                cco: $('.cco').serializeArray().map(function (m) { return m.value }),
                correo: $('#correo1').val(),
                user: localStorage.getItem('user') 
            },
            success: function (result) {
                if (result.done) {
                    alertify.success(result.message)
                    window.location.reload();
                } else {
                    alertify.error(result.message)
                }
            }
        })
        return false;

    });

});

function addPara() {
    var div = $('<div>', { class: 'form-group' })
        .append($('<div>', { class: 'input-group' })
            .append($('<input>', { type: 'email', name: 'para', class: 'form-control para' }))
            .append($('<span>', {
                class: 'input-group-addon btn btn-danger',
                click: function () {
                    div.remove();
                }
            })
                .append($('<i>', { class: 'fa fa-trash' }))))
        .insertAfter($('#para-fg'))
}

function addCC() {
    var div = $('<div>', { class: 'form-group' })
        .append($('<div>', { class: 'input-group' })
            .append($('<input>', { type: 'email', name: 'cc', class: 'form-control cc' }))
            .append($('<span>', {
                class: 'input-group-addon btn btn-danger',
                click: function () {
                    div.remove();
                }
            })
                .append($('<i>', { class: 'fa fa-trash' }))))
        .insertAfter($('#cc-fg'))
}

function addCCO() {
    var div = $('<div>', { class: 'form-group' })
        .append($('<div>', { class: 'input-group' })
            .append($('<input>', { type: 'email', name: 'cco', class: 'form-control cco' }))
            .append($('<span>', {
                class: 'input-group-addon btn btn-danger',
                click: function () {
                    div.remove();
                }
            })
                .append($('<i>', { class: 'fa fa-trash' }))))
        .insertAfter($('#cco-fg'))
}


function createInitialView(quotation) {
    $.ajax({
        url: 'handlers/DetalleCotizacion.ashx',
        type: 'POST',
        data: { 1: 'getInitialData', quotationId: quotation },
        success: function (result) {
            if (result.done) {

                COTIZACION = result.data;
                var moneda = result.data.quotationMonedaId;
                $('#titleNumeroCotizacion').text('COTIZACION IT ' + result.data.quotatationIT);

                //CLIENT DATA
                $('#rc_rut').text(result.data.clientRut);
                $('#rc_nombre').text(result.data.clientName);
                $('#rc_direccion').text(result.data.clientDirection);
                $('#rc_fono').text(result.data.clientPhone);
                $('#rc_email').text(result.data.clientMail);
                $('#rc_giro').text(result.data.clientGiro);
                $('#rc_ciudad').text(result.data.clientCiudad);
                $('#rc_contacto').text(result.data.clientContactName);

                // QUOTATION DATA
                $('#rg_fecha_doc').text(result.data.quotationFechaDoc);
                $('#rg_fuente_solicitud').text(result.data.quotationFuenteSolicictud);
                $('#rg_forma_de_pago').text(result.data.quotationFormaPago);
                $('#rg_moneda').text(result.data.quotationMoneda);
                $('#rg_vendedor').text(result.data.quotationVendedor);
                $('#rg_validez').text(result.data.quotationValidez);
                $('#rg_observaciones').text(result.data.quotationObservation);
                $('#rg_nota').text(result.data.quotationNota);
                $('#rg_estado').text(result.data.quotationEstado);

                $('#r_valor').text(result.data.quotationValor.toLocaleString('es-CL', { maximumFractionDigits: moneda == 1 ? 0 : 2, minimumFractionDigits: 0 }));
                $('#r_descuento').text(result.data.quotationDescuento.toLocaleString('es-CL', { maximumFractionDigits: moneda == 1 ? 0 : 2, minimumFractionDigits: 0 }));
                $('#r_recargo').text(result.data.quotationRecargo.toLocaleString('es-CL', { maximumFractionDigits: moneda == 1 ? 0 : 2, minimumFractionDigits: 0 }));
                $('#r_total').text(result.data.quotationTotal.toLocaleString('es-CL', { maximumFractionDigits: moneda == 1 ? 0 : 2, minimumFractionDigits: 0 }));

                MailCliente = result.data.clientMail;

                //Solo Muestra boton envia eMail de Seguimiento de cotizacion si esta en estado (1,2) 
                //Y si las todas las inspecciones estan en estado (2), tambien puede enviar eMail
                if (result.data.quotationStateId != 2) {
                    $('#send-mail').hide('fast');
                }

                //Si el estado es (1,2) manda correo de Seguimeinto
                if (result.data.quotationStateId == 1 || result.data.quotationStateId == 2) {
                    GetMailSeguimientoCotizacion(quotation);
                }

                //Si las inspecciones terminan, manda correo de Post Venta
                if (result.data.inspeccionesCount) {
                    GetMailPostVenta();
                }

                console.log(result.data)
            }
            else {
                alertify.error(result.message);
            }

        }
    })
}

function reloadGridHitos(quotationId) {
    jQuery("#grid")
        .jqGrid('setGridParam',
        {
            url: 'handlers/DetalleCotizacion.ashx',
            postData: {
                1: 'grid',
                id: quotationId
            },
        })
        .trigger("reloadGrid");
}

function createItemsView(quotation) {
    var moneda;
    var TotalCantidad = 0;
    var TotalUnitario = 0;
    var TotalDescuento = 0;
    var TotalValorCliente = 0;
    $.ajax({
        url: 'handlers/DetalleCotizacion.ashx',
        type: 'POST',
        data: { 1: 'quotationItems', quotationId: quotation },
        success: function (result) {
            if (result.done) {
                var tbody, trthead, trthead2;
                $('#grid-items-resume')
                    .append($('<thead>')
                        .append(trthead = $('<tr>')))
                    .append(tbody = $('<tbody>'));

                $(result.data).each(function (h, data) {
                    moneda = parseInt(data.moneda);
                    
                    tbody
                        .append($('<tr>')
                            .append($('<td>')
                                .addClass('text-right')
                                .text(data.cantidad))
                            .append($('<td>')
                                .text(data.descripcionEditada))
                            .append($('<td>')
                                .text(data.ubicacion))
                            .append($('<td>')
                                .addClass('text-right')
                                .text(data.valorUnitario.toLocaleString('es-CL', { maximumFractionDigits: moneda == 1 ? 0 : 2, minimumFractionDigits: 0 })))
                            .append($('<td>')
                                .text(data.modoDescuento))
                            .append($('<td>')
                                .addClass('text-right')
                                .text(data.valorDescuento.toLocaleString('es-CL', { maximumFractionDigits: moneda == 1 ? 0 : 2, minimumFractionDigits: 0 })))
                            .append($('<td>')
                                .addClass('text-right')
                                .text(data.precioCliente.toLocaleString('es-CL', { maximumFractionDigits: moneda == 1 ? 0 : 2, minimumFractionDigits: 0 })))
                    );

                    TotalCantidad += data.cantidad;
                    TotalUnitario += data.valorUnitario;
                    TotalDescuento += data.valorDescuento;
                    TotalValorCliente += data.precioCliente;

                })


                tbody
                    .append($('<tr>')
                        .append($('<td>')
                            .addClass('text-right')
                            .addClass('negrita')
                            .text(TotalCantidad))
                        .append($('<td>')
                            .text(''))
                        .append($('<td>')
                            .text(''))
                        .append($('<td>')
                            .addClass('text-right')
                            .addClass('negrita')
                            .text(TotalUnitario.toLocaleString('es-CL', { maximumFractionDigits: moneda == 1 ? 0 : 2, minimumFractionDigits: 0 })))
                        .append($('<td>')
                            .text(''))
                        .append($('<td>')
                            .addClass('text-right')
                            .addClass('negrita')
                            .text(TotalDescuento.toLocaleString('es-CL', { maximumFractionDigits: moneda == 1 ? 0 : 2, minimumFractionDigits: 0 })))
                        .append($('<td>')
                            .addClass('text-right')
                            .addClass('negrita')
                            .text(TotalValorCliente.toLocaleString('es-CL', { maximumFractionDigits: moneda == 1 ? 0 : 2, minimumFractionDigits: 0 })))
                    );


                console.log(result.data)
            }
            else {
                alertify.error(result.message);
            }

        }
    })
}

function createCostosLogisticosView(quotation) {
    var Total = 0;
    var moneda;
    $.ajax({
        url: 'handlers/DetalleCotizacion.ashx',
        type: 'POST',
        data: { 1: 'GetCostosLogisticos', quotationId: quotation },
        success: function (result) {
            if (result.done) {
                var tbody, trthead, trthead2;
                $('#gastos-table-resume')
                    .append($('<thead>')
                        .append(trthead = $('<tr>')))
                    .append(tbody = $('<tbody>'));

                $(result.data).each(function (h, data) {
                    moneda = parseInt(data.moneda);
                    tbody
                        .append($('<tr>')
                            .append($('<td>')
                                .addClass('text-right')
                                .text(data.cantidad))
                            .append($('<td>')
                                .text(data.tipoGasto))
                            .append($('<td>')
                                .text(data.valor.toLocaleString('es-CL', { maximumFractionDigits: moneda == 1 ? 0 : 2, minimumFractionDigits: 0 })))
                            .append($('<td>')
                                .addClass('text-right')
                                .text(data.total.toLocaleString('es-CL', { maximumFractionDigits: moneda == 1 ? 0 : 2, minimumFractionDigits: 0 })))

                        );
                    Total += data.total;
                })


                $('#total-gastos-resume').addClass('text-right').text(Total.toLocaleString('es-CL', { maximumFractionDigits: moneda == 1 ? 0 : 2, minimumFractionDigits: 0 }));

                console.log(result.data)
            }
            else {
                alertify.error(result.message);
            }

        }
    })
}

function GetMailSeguimientoCotizacion(c) {
    $.ajax({
        url: 'handlers/Cotizaciones.ashx',
        data: {
            1: 'get-global'
        },
        success: function (result) {
            if (result.done) {
                GLOBAL = result.data;
                $.ajax({
                    url: 'handlers/DetalleCotizacion.ashx',
                    data: {
                        1: 'getmailFormat',
                        quotationId: c,
                        mail: GLOBAL.mail_seguimiento_cotizacion
                    },
                    success: function (result) {
                        if (result.done) {
                            $('#correo').jqteVal(result.mail);
                        }
                    }
                })

            }
        }
    })
}

function GetMailPostVenta() {
    $.ajax({
        url: 'handlers/Cotizaciones.ashx',
        data: {
            1: 'get-global'
        },
        success: function (result) {
            if (result.done) {
                GLOBAL = result.data;
                $('#correo').jqteVal(GLOBAL.mail_post_venta);
            }
        }
    })
}

function getGlobal() {
    $.ajax({
        url: 'handlers/Cotizaciones.ashx',
        data: {
            1: 'get-global'
        },
        success: function (result) {
            if (result.done) {
                GLOBAL = result.data
                $('#asunto').val(result.data.mail_envio_cotizacion_subject);
                $('#correo1').jqteVal(result.data.mail_envio_cotizacion_text);
            }
        }
    })
}