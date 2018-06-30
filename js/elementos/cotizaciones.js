
var URL_HANDLER = 'handlers/Cotizaciones.ashx'
var SELECTED_CLIENT
var SELECTED_PRODUCT
var SELECTED_COTIZACION
var IVA = 0.19;
var GLOBAL;
var CLOSE_DETAIL_DIALOG = false;
var MailCliente = ''
$(document).ready(function () {

    $(window).on('beforeunload', function () {
        return '¿Estás seguro que quieres salir?';
    });
    getGlobal();
    $('#fecha-fase2').datetimepicker({
        locale: 'es',
        format: 'DD-MM-YYYY HH:mm'
    });
    $('#uf').text(localStorage.getItem('UF'))
    $('#dolar').text(localStorage.getItem('Dolar'))
    $('#euro').text(localStorage.getItem('Euro'))
    $('#correo, #correo2, #f2-correo').jqte();
    // RUT
    $('#rut-cliente').Rut({
        format_on: 'keyup',
        validation: true,
        on_error: function () {
            alertify.error('El Rut ingresado no corresponde')
            $('#rut-cliente').focus()
        }
    })
    $('#fc_rut').Rut({ format_on: 'keyup' })
    $('#rut-cliente').blur(function () {
        getClientByRut();
    })
    // Combos
    combobox($('#forma_pago'), { 1: 'formasDePago' }, '');
    combobox($('#moneda'), { 1: 'monedas' }, '');
    combobox($('#vendedor'), { 1: 'vendedores' }, 'Seleccione...');
    combobox($('#f_estado'), { 1: 'estadosCotizacion' }, 'Todos');
    combobox($('#tipo_elevador'), { 1: 'aparatos' }, ' ')
    combobox($('#tipo_funcionamiento'), { 1: 'tipoFuncionamiento' }, ' ')
    combobox($('#marca'), { 1: 'marcas' }, ' ')
    combobox($('#uso'), { 1: 'destinoProyecto' }, ' ')
    combobox($('#tipo-gasto'), { 1: 'tipoGasto' }, 'Tipo Gasto')
    combobox($('#producto'), { 1: 'productos' }, '');
    combobox($('#ciudad'), { 1: 'ciudades' }, '');
    loadCmbAnoInstalacion();

    // Wizard
    $('#rootwizard').bootstrapWizard({
        onTabChange: function (ul, li, index, index2) {
            console.log('index', index2)
            if (index === 0 && index2 >= 1 && index2 < 5) {
                if (!validateForm($('#form-cliente'))) {
                    alertify.error('Complete los datos requeridos')
                    return false
                } else {
                    $('#para').val($('#email-contacto-cliente').val())
                }
            }
            if (index === 1 && index2 > 1) {
                if (!validateForm($('#form-generales'))) {
                    alertify.error('Complete los datos requeridos2')
                    return false
                }

            }
            if (index === 2 && index2 > 2) {
                if (!ITEMS || !ITEMS.length || ITEMS.length == 0) {
                    alertify.error('Ingrese al menos 1 ítem')
                    return false;
                }
            }
            if (index2 == 4) {

                getResume()
            }
        }
    });
    $('#rootwizard .finish').click(function () {
        if (!validateForm($('#form-cliente'))) {
            alertify.error('Complete los datos cliente')
            return false
        }
        if (!validateForm($('#form-generales'))) {
            alertify.error('Complete los datos generales')
            return false
        }
        if (!ITEMS || !ITEMS.length || ITEMS.length == 0) {
            alertify.error('Ingrese al menos 1 ítem')
            return false;
        }
        alertify.confirm('GUARDAR COTIZACIÓN', '¿Está seguro/a de guardar la cotización?', function () {
            saveAll()
        }, function () { })
    });
    // Datepicker
    var today = new Date()
    var day = today.getDate()
    var month = today.getMonth() + 1
    var year = today.getFullYear()
    var date = (day < 10 ? '0' + day : day) + '-' + (month < 10 ? '0' + month : month) + '-' + year;
    $('#fecha_doc').datepicker({ dateFormat: 'dd-mm-yy' }).val(date)
    $('#validez').datepicker({ dateFormat: 'dd-mm-yy', minDate: 0 })

    $('#item_tipo_descuento').change(function () {
        var val = $(this).val();
        if (val == '0') {
            $('#item_descuento').val(0).prop('disabled', true)
        } else if (val == 'VALOR') {
            $('#item_descuento').prop({ disabled: false, max: 1000000000 })
        } else if (val == 'PORCENTAJE') {
            $('#item_descuento').val(0).prop({ disabled: false, max: 100 })

        }
    })
    $('#tipo-gasto').change(function () {
        setPriceLogisticCoste($(this).val())
    })
    $('#f_inicio, #f_fin').datepicker({ dateFormat: 'dd-mm-yy' })

    // Grid Cotizaciones
    $('#grid').jqGrid({
        url: URL_HANDLER,
        postData: {
            1: 'grid-cotizaciones',
            cliente: '',
            estado: 0,
            inicio: '',
            fin: '',
            alertF2: false
        },
        mtype: "POST",
        autowidth: true,
        datatype: "json",
        shrinkToFix: true,
        ajaxGridOptions: { global: false },
        colModel: [
            { label: 'Folio', name: 'Id', index: 'Id', key: true, hidden: true, width: 10, align: 'center', sortable: true },
            { label: 'IT', name: 'IT', index: 'Id', width: 20, align: 'center', sortable: true },
            { label: 'ClienteId', name: 'ClienteId', hidden: true },
            { label: 'Cliente', name: 'Cliente', index: 'Cliente.Nombre', width: 50, align: 'left', sortable: false },
            { label: 'Fecha Doc', name: 'FechaDoc', index: 'FechaDoc', width: 20, align: 'center', sortable: true },
            { label: 'Nombre Contacto', name: 'NombreContacto', index: 'Cliente.NombreContacto', width: 30, align: 'left', sortable: true },
            { label: 'Responsable', name: 'Responsable', index: 'Usuario1.Nombre', width: 25, align: 'left', sortable: true },
            { label: 'Estado', name: 'Estado', index: 'EstadoCotizacion.Descripcion', width: 15, align: 'left', sortable: true },
            { label: 'Moneda', name: 'Moneda', index: 'Moneda.Descripcion', width: 15, align: 'left', sortable: true },
            { label: 'Observacion', name: 'Observacion', index: 'Observacion', width: 30, align: 'left', sortable: true, hidden: true },
            { label: 'Fecha Creación', name: 'FechaCreacion', index: 'FechaCreacion', width: 20, align: 'center', sortable: true },
            { label: 'Valor', name: 'Valor', index: 'Valor', width: 15, align: 'right', sortable: true, formatter: 'number', hidden: true, formatoptions: { decimalSeparator: ",", thousandsSeparator: ".", decimalPlaces: 2 } },
            { label: 'Descuento', name: 'Descuento', index: 'Descuento', width: 15, align: 'right', sortable: true, formatter: 'number', hidden: true, formatoptions: { decimalSeparator: ",", thousandsSeparator: ".", decimalPlaces: 2 } },
            { label: 'Recargo', name: 'Recargo', index: 'Recargo', width: 15, align: 'right', sortable: true, formatter: 'number', hidden: true, formatoptions: { decimalSeparator: ",", thousandsSeparator: ".", decimalPlaces: 2 } },
            { label: 'Total', name: 'Total', index: 'Total', width: 15, align: 'right', sortable: true, formatter: 'number', formatoptions: { decimalSeparator: ",", thousandsSeparator: ".", decimalPlaces: 2 } },
            { label: 'PDF', width: 10, formatter: pdfButton, align: 'center' },
            { label: 'Ir', width: 10, formatter: detailButton, align: 'center' },
            { label: 'Servicio', name: 'HasService', width: 10, formatter: serviceButton, align: 'center' },
            { label: 'Fase 2', name: 'AlertFase2', width: 10, formatter: alertfase2, align: 'center' },
            { label: '', name: 'EstadoId', hidden: true },
            { label: '', name: 'MailCliente', hidden: true },
            { label: '', name: 'AlertaFase2Enviada', hidden: true },
            { label: '', name: 'HasFase2', hidden: true },
            { label: '', name: 'IsOk', hidden: true },
            { label: 'Mail', name: 'SentMail', width: 10, formatter: mailButton, align: 'center'}
        ],
        sortname: 'Id',
        sortorder: 'desc',
        viewrecords: true,
        height: 320,
        rowNum: 30,
        pager: "#pager",
        caption: 'Cotizaciones registrados',
        onCellSelect: function (rowid, iCol, cellcontent, e) {
            var rowData = $(this).jqGrid('getRowData', rowid);
            SELECTED_COTIZACION = rowData;
            MailCliente = rowData.MailCliente;
            if (iCol == 15) {
                getPdf(rowData.Id)
            }
            if (iCol == 16) {
                window.open('DetalleCotizacion.aspx?c=' + rowData.Id, '_blank', 'fullscreen=yes')
            }
            if (iCol == 17) {
                if (rowData.HasService != '') {
                    openServiceDetail(rowData);
                }
            }
            if (iCol == 18) {
                if (rowData.AlertFase2 != '') {
                    openSendAlertFase2()
                }
            }
            if (iCol == 24) {
                sendMail();
            }

        }
    }).navGrid('#pager',
        {
            edit: false, add: false, del: false, search: false,
            refresh: true, view: false, position: "left", cloneToTop: false
        });
    $("#grid").navButtonAdd('#pager', {
        caption: "Exportar",
        title: "Exportar cotizaciones a Excel",
        buttonicon: "ui-icon-document",
        onClickButton: function () {
            window.location.href = 'handlers/ExportExcel.ashx?1=cotizaciones&cliente=' + $('#f_name').val() + '&estado=' + $('#f_estado').val() + '&inicio=' + $('#f_inicio').val() + '&fin=' + $('#f_fin').val()
        },
        position: "first"
    });

    $('#services-grid').jqGrid({
        mtype: "POST",
        autowidth: true,
        datatype: "json",
        shrinkToFix: true,
        ajaxGridOptions: { global: false },
        colModel: [
            { label: 'Id', name: 'Id', key: true, hidden: true },
            { label: 'IT', name: 'It', index: 'IT', width: 45, align: 'left', sortable: false },
            { label: 'Fase', name: 'Fase', index: 'Fase', width: 20, align: 'center', sortable: true },
            { label: 'Calificación', name: 'Calificacion', index: 'Calificacion', width: 80, formatter: calificacionLabel, width: 50, align: 'center', sortable: true },
            { label: 'Calificación', name: 'Calificacion1', index: 'Calificacion', hidden: true }
        ],
        sortname: 'IT',
        sortorder: 'asc',
        viewrecords: true,
        height: 200,
        rowNum: 50,
        pager: "#services-pager",
        caption: 'Servicios relacionados',
        width: 'auto',
        onCellSelect: function (rowid, iCol, cellcontent, e) {
            
        },
        loadComplete: function (data) {
            if (!data.rows)
                return;
            console.log(data.rows)
            chartQualifyServices(data.rows)
        }
    }).navGrid('#services-pager',
        {
            edit: false, add: false, del: false, search: false,
            refresh: true, view: false, position: "left", cloneToTop: false
        });


    $('#grid-clientes').jqGrid({
        url: URL_HANDLER,
        postData: {
            1: 'grid-clientes',
            name: $('#fc_nombre').val(),
            rut: $('#fc_rut').val()
        },
        mtype: "POST",
        autowidth: true,
        datatype: "json",
        shrinkToFix: true,
        ajaxGridOptions: { global: false },
        colModel: [
            { label: 'Id', name: 'Id', key: true, hidden: true },
            { label: 'RUT', name: 'Rut', index: 'Rut', width: 15, align: 'left', sortable: false },
            { label: 'Nombre', name: 'Nombre', index: 'Nombre', width: 30, align: 'left', sortable: true },
            { label: 'Dirección', name: 'Direccion', index: 'Direccion', width: 35, align: 'left', sortable: true },
            { label: 'Teléfono', name: 'Telefono', index: 'Telefono', width: 30, align: 'left', sortable: true },
            { label: 'Giro', name: 'Giro', index: 'Giro', width: 30, align: 'left', sortable: true },
            { label: 'Nombre Contacto', name: 'NombreContacto', index: 'NombreContacto', width: 30, align: 'left', sortable: true },
            { label: 'E-mail', name: 'Email', index: 'Email', width: 30, align: 'left', sortable: true },
            { label: 'Seleccionar', width: 5, formatter: selectButton, align: 'center' }
        ],
        sortname: 'Nombre',
        sortorder: 'asc',
        viewrecords: true,
        height: 350,
        rowNum: 30,
        pager: "#pager-clientes",
        caption: 'Clientes registrados',
        onCellSelect: function (rowid, iCol, cellcontent, e) {
            if (iCol == 8) {
                SELECTED_CLIENT = $(this).getRowData(rowid)
                setSelectedClient()
                $('#clientes-dialog').dialog('close')
            }

        }
    }).navGrid('#pager-clientes',
        {
            edit: false, add: false, del: false, search: false,
            refresh: true, view: false, position: "left", cloneToTop: false
        });
    function selectButton() {
        return '<div class="btn-grid"><i class="fa fa-check"></i></div>'
    }
    function pdfButton() {
        return '<div class="btn-grid"><i class="fa fa-file-pdf-o" style="color: #F44336"></i></div>'
    }
    function detailButton() {
        return '<div class="btn-grid"><i class="fa fa-arrow-right" style="color: #3F51B5"></i></div>'
    }
    function serviceButton(cellvalue, options, rowObject) {
        if(cellvalue)
            return '<div class="btn-grid"><i class="fa fa-list" style="color: #64DD17"></i></div>'
        return ''
    }
    function alertfase2(cellvalue, options, rowObject) {
        var div = '';
        var sent = rowObject.AlertaFase2Enviada;
        var color = sent ? '#4CAF50' : '#F44336';
        var clase = sent ? 'fa fa-envelope-o' : 'fa fa-envelope-o fa-pulse'
        var title = sent ? 'Alerta ya enviada. ¿Quieres enviarla nuevamente?' : 'Enviar alerta'
        if (cellvalue)
            div = '<div class="btn-grid" title="' + title + '"><i class="' + clase + '" style="color: ' + color + '"></i></div>'
        return div;
    }
    function mailButton(cellvalue, options, rowObject) {
        var color = cellvalue ? 'green' : 'red'
        return '<div class="btn-grid" title="Enviar E-mail"><i class="fa fa-envelope" style="color: ' + color +  '"></i></div>'
    }
    function calificacionLabel(cellvalue, options, rowObject) {
        var label = '';
        switch (cellvalue)
        {
            case null:
            case undefined:
            case '':
                label = '<h5><label class="label label-default">No se ha ingresado calificación</label></h5>'; break;
            case 1: label = '<h5><label class="label label-success">Califica</label></h5>'; break;
            case 2: label = '<h5><label class="label label-warning">Califica con observaciones</label></h5>'; break;
            case 0: label = '<h5><label class="label label-danger">No califica</label></h5>'; break;
        }
        return label;
        
    }
   
    $('#add-dialog').dialog({
        modal: true, bgiframe: false, width: '100%', title: 'Nueva Cotización', draggable: true,
        resizable: false, closeOnEscape: true, autoOpen: false, show: 'clip', hide: 'puff',
        position: 'top',
        open: function (type, data) {
            $('#rootwizard').find("a[href*='tab5']").trigger('click');
            $('#rootwizard').find("a[href*='tab1']").trigger('click');
            reloadGridClientes()
            //reloadGridProductos()
        },
        close: function () {
            resetAll();
        }
    });
    $('#add-item-dialog').dialog({
        modal: true, bgiframe: false, width: '90%', title: 'Nuevo Item', draggable: true,
        resizable: false, closeOnEscape: true, autoOpen: false, show: 'clip', hide: 'puff',
        open: function (type, data) {
            $('#ubicacion').val(SELECTED_CLIENT.Direccion);
            $('#item_descripcion').prop('readonly', true)
        },
        close: function () {
            $('#form-detalle')[0].reset()
        },
        buttons: [
            {
                id: 'add-item',
                text: 'Agregar y Cerrar',
                click: function () {
                    CLOSE_DETAIL_DIALOG = true;
                    $('#form-detalle').submit();
                }
            },
            {
                id: 'add-item2',
                text: 'Agregar y Mantener',
                click: function () {
                    CLOSE_DETAIL_DIALOG = false;
                    $('#form-detalle').submit();
                }
            },
            {
                id: 'add-item-close',
                text: 'Cerrar',
                click: function () {
                    $(this).dialog('close');
                }
            }
        ]
    });
    $('#add-file-dialog').dialog({
        modal: true, bgiframe: false, width: 400, title: 'Nuevo Archivo', draggable: true,
        resizable: false, closeOnEscape: true, autoOpen: false, show: 'clip', hide: 'puff',
        open: function (type, data) {

        },
        close: function () {
            $('#add-file-form')[0].reset()
        },
        buttons: [
            {
                text: 'Guardar',
                click: function () {
                    $('#add-file-form').submit();
                }
            }
        ]
    })
    $('#post-venta').click(function () {

        $('#send-mail-dialog').dialog({
            modal: true, bgiframe: false, width: 700, title: 'Enviar correo de post-venta', draggable: true,
            resizable: false, closeOnEscape: true, autoOpen: true, show: 'clip', hide: 'puff',
            open: function (type, data) {
                $(this).find('form')[0].reset();
                $('#mail-destinatario').val(MailCliente);
                $('#correo2').jqteVal(GLOBAL.mail_post_venta);
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
        sendMailPosVenta();
        return false;

    });
    $('#send-mail-cot-form').submit(function (e) {
        sendMailCotizacion();
        return false;

    });
    $('#services-dialog').dialog({
        modal: true, bgiframe: false, width: 600, title: 'Servicios', draggable: true,
        resizable: false, closeOnEscape: true, autoOpen: false, show: 'clip', hide: 'puff',
        position: { my: "center", at: "top", of: document, collision: "fit" },
        open: function (type, data) {
            $('#agendar-fase2').hide();
            $('#post-venta').hide();
            if(SELECTED_COTIZACION.IsOk == 'true' && SELECTED_COTIZACION.EstadoId != '5') {
                $('#post-venta').show();
            }
            if (SELECTED_COTIZACION.HasFase2 != 'true') {
                $('#agendar-fase2').show();
            }
        },
        close: function () {
            
        },
        buttons: [
            
        ]
    })
    $('#clientes-dialog').dialog({
        modal: true, bgiframe: false, width: '90%', title: 'Buscar cliente', draggable: true,
        resizable: false, closeOnEscape: true, autoOpen: false, show: 'clip', hide: 'puff'
    });
    $('#productos-dialog').dialog({
        modal: true, bgiframe: false, width: '90%', title: 'Buscar producto', draggable: true,
        resizable: false, closeOnEscape: true, autoOpen: false, show: 'clip', hide: 'puff'
    });
    $('#add-select-dialog').dialog({
        modal: true, bgiframe: false, width: '50%', title: 'Seleccione cómo desea comenzar', draggable: true,
        resizable: false, closeOnEscape: true, autoOpen: false, show: 'clip', hide: 'puff',
        close: function () {
            $('#a-partir-de-it').val('');
        }
    });
    $('#add').click(function () {
        //$('#add-dialog').dialog('open')
        $('#add-select-dialog').dialog('open')
    })
    $('#desde-cero').click(function () {
        $('#add-dialog').dialog('open');
        $('#add-select-dialog').dialog('close')
    })
    $('#a-partir-de').click(function () {
        aPartirDe();
    })
    $('#a-partir-de-it').keyup(function (e) {
        if(e.keyCode == 13)
            aPartirDe();
    })
    $('#search-client').click(function () {
        $('#clientes-dialog').dialog('open')
    })
    $('#search_product').click(function () {
        $('#productos-dialog').dialog('open')
    })
    $('#add-item-button').click(function () {
        $('#add-item-dialog').dialog('open')
    })
    $('#add-gasto').click(function () {
        $('#gastos-form').submit();
    })
    $('#f_remove').click(function () {
        $('#formFiltros')[0].reset();
    })
    $('#f_name').keyup(function () {
        $('#formFiltros').submit()
    })
    $('#f_inicio, #f_fin').change(function () {
        $('#formFiltros').submit()
    })
    $('#f_estado').change(function () {
        $('#formFiltros').submit()
    })
    $('#formFiltros').submit(function () {
        reloadGridCotizaciones()
        return false;
    })
    $('#gastos-form').submit(function () {
        if (!validateForm($(this))) {
            alertify.error('Complete los datos requeridos');
            return false;
        }
        var valorGasto = parseFloat($('#valor-gasto').val()) || 0;
        if (valorGasto == 0) {
            alertify.error('Debe incluir un valor');
            return false;
        }
        var cant = $('#gastos-cantidad').val();
        if (!$.isNumeric(cant) || parseInt(cant) < 0) {
            alertify.error('Ingrese una cantidad válida');
            return false;
        }
        var q = parseInt(cant);
        var total = q * valorGasto
        addGasto(q, $('#tipo-gasto').val(), $('#tipo-gasto option:selected').text(), valorGasto, total);
        return false;
    })
    $('#form-detalle').submit(function () {
        if (!validateForm($(this))) {
            alertify.error('Complete los datos requeridos');
            return false;
        }

        var moneda = parseInt($('#moneda').val())

        var unitario = parseFloat($('#item_unitario').val());
        var mododescuento = $('#item_tipo_descuento').val();
        var descuento = parseFloat($('#item_descuento').val()).toFixed(2);
        var cantidad = parseInt($('#item_cantidad').val());
        var tipo_elevador = parseInt($('#tipo_elevador').val());
        var tipo_funcionamiento = parseInt($('#tipo_funcionamiento').val())
        var marca = parseInt($('#marca').val())
        var ano_instalacion = $('#ano_instalacion').val()
        var empresa_instaladora = $('#empresa_instaladora').val()
        var ubicacion = $('#ubicacion').val()
        var tipo_elevador_name = $('#tipo_elevador option:selected').text()
        var tipo_funcionamiento_name = $('#tipo_funcionamiento option:selected').text()
        var marca_name = $('#marca option:selected').text();
        var uso = $('#uso').val();
        var altura = $('#altura').val();

        if (mododescuento == 'PORCENTAJE' && descuento > 100) {
            alertify.error('El porcentaje no puede ser mayor a 100');
            return false;
        }

        var desc = mododescuento == 'PORCENTAJE'
            ? (parseFloat(unitario) * parseFloat(descuento) / 100)
            : mododescuento == 'VALOR'
                ? parseFloat(descuento) : 0;
        var m = mododescuento == 'PORCENTAJE' ? '%' : '$'
        var precioCliente = ((unitario - desc) * cantidad);
        var item = {
            Id: Math.random().toString(36).slice(2),
            Producto: $('#producto').val(),
            Descripcion: $('#item_descripcion').val(),
            Cantidad: cantidad,
            ValorUnitario: unitario,
            ModoDescuento: m,
            Descuento: descuento,
            ValorDescuento: desc,
            PrecioCliente: precioCliente,
            Exento: true,
            TipoElevador: tipo_elevador,
            TipoElevadorName: tipo_elevador_name,
            TipoFuncionamientoName: tipo_funcionamiento_name,
            MarcaName: marca_name,
            TipoFuncionamiento: tipo_funcionamiento,
            Marca: marca,
            AnoInstalacion: ano_instalacion,
            EmpresaInstaladora: empresa_instaladora,
            Ubicacion: ubicacion,
            Uso: uso,
            Altura: altura
        }
        addItem(item)
        return false;
    })
    $('#add-file-form').submit(function () {
        var formData = new FormData();
        formData.append('1', 'uploadFile');
        formData.append('file', $('#add-file-file')[0].files[0]);
        formData.append('name', $('#add-file-name').val());
        formData.append('uso', 'NUEVA_COTIZACION_MAIL');
        $.ajax({
            url: URL_HANDLER,
            data: formData,
            type: 'post',
            dataType: 'json',
            success: function (result) {
                if (result.done) {
                    alertify.success(result.message);
                    $('#add-file-dialog').dialog('close');
                    getFiles();
                }
                else
                    alertify.error(result.message);
            },
            processData: false,
            contentType: false,
            error: function () {
                alertify.error('Error');
            }
        });
        return false;
    });

    $('#send-alert-fase2-form').submit(function () {
        sendAlertFase2()
        return false;
    })

    $('#fc_nombre, #fc_rut').keyup(function () {
        reloadGridClientes();
    })
    $('#recargo').blur(function () {
        updateValues();
    })
    $('#no_send_mail').change(function () {
        var isChecked = $(this).is(':checked')
        if (isChecked) {
            $('#para').val('').prop('disabled', true)
            $('#asunto').val('').prop('disabled', true);
            $('#correo').jqteVal('').prop('disabled', true);
            $('.chk-file').prop('disabled', true);
        } else {
            $('#para').val(SELECTED_CLIENT.Email).prop('disabled', false);
            $('#asunto').val(GLOBAL.mail_envio_cotizacion_subject).prop('disabled', false);
            $('#correo').jqteVal(GLOBAL.mail_envio_cotizacion_text).prop('disabled', false);
            $('.chk-file').prop('disabled', false);
        }
    })
    $('#edit-item-description').click(function () {
        $('#item_descripcion').prop('readonly', !$('#item_descripcion').prop('readonly')).focus();
    }).tooltip()
    $('#edit-price').click(function () {
        $('#item_unitario').prop('readonly', !$('#item_unitario').prop('readonly')).focus();

    })
    $('#item_cantidad, #tipo_elevador, #tipo_funcionamiento, #marca, #ano_instalacion, #empresa_instaladora, #altura, #uso, #producto, #ubicacion').change(function () {
        if ($('#item_descripcion').prop('readonly'))
            $('#item_descripcion').val(setDescripcion());
        if ($('#item_unitario').prop('readonly'))
            getPrice();
    })
    $('#edit-valor-gasto').click(function () {
        $('#valor-gasto').prop('readonly', $('#tipo-gasto').val()  == 0 ? true : !$('#valor-gasto').prop('readonly')).focus();
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
    $('#item_descuento').blur(function () {
        setPrice();
    })
    $('#add-file').click(function () {
        $('#add-file-dialog').dialog('open')
    })
    $('#agendar-fase2').click(function () {
        openAgendaFase2();
    })
    getFiles();
})

function sendMail() {
    $('#send-mail-cot-dialog').dialog({
        modal: true, bgiframe: false, width: 700, title: 'Enviar correo', draggable: true,
        resizable: false, closeOnEscape: true, autoOpen: true, show: 'clip', hide: 'puff',
        open: function (type, data) {
            $(this).find('form')[0].reset();
            $('#para').val(MailCliente);
            $('#correo').jqteVal(GLOBAL.mail_envio_cotizacion_text);
            $('#asunto').val(GLOBAL.mail_envio_cotizacion_subject);
        },
        close: function () {

        },
        buttons: [
            {
                id: 'btn-send-mail',
                text: 'Enviar',
                click: function () {
                    $('#send-mail-cot-form').submit();
                }
            }
        ]
    });
}
function sendAlertFase2() {
    $.ajax({
        url: URL_HANDLER,
        type: 'POST',
        data: {
            1: 'sendAlertaFase2',
            id: SELECTED_COTIZACION.Id,
            to: $('#f2-destinatario').val(),
            subject: $('#f2-asunto').val(),
            mail: $('#f2-correo').val(),
            user: localStorage.getItem('user')
        },
        success: function (result) {
            if (result.done) {
                alertify.success(result.message)
                $('#send-alert-fase2-dialog').dialog('close')
                reloadGridCotizaciones();
            } else {
                alertify.error(result.message)
            }
        }
    })
}
function completingSendAlertFase2() {
    $.ajax({
        url: URL_HANDLER,
        type: 'POST',
        data: {
            1: 'completingSendAlertaFase2',
            id: SELECTED_COTIZACION.Id
        },
        success: function (result) {
            if (result.done) {
                $('#f2-correo').jqteVal(result.mail)
                $('#f2-destinatario').val(result.direccion)
            } else {
                alertify.error(result.message)
            }
        }
    })
}
function openSendAlertFase2() {
    $('#send-alert-fase2-dialog').dialog({
        modal: true, bgiframe: false, width: 700, title: 'Enviar correo de alerta de fase 2', draggable: true,
        resizable: false, closeOnEscape: true, autoOpen: true, show: 'clip', hide: 'puff',
        open: function (type, data) {
            $(this).find('form')[0].reset();
            completingSendAlertFase2()
        },
        close: function () {

        },
        buttons: [
            {
                id: 'btn-send1',
                text: 'Enviar',
                click: function () {
                    $('#send-alert-fase2-form').submit();
                }
            }
        ]
    });
}

function openAgendaFase2() {
    $('#fase2-dialog').dialog({
        modal: true, bgiframe: false, width: 600, height: 500, title: 'Crear Fase 2', draggable: true,
        resizable: false, closeOnEscape: true, autoOpen: true, show: 'clip', hide: 'puff',
        position: { my: "center", at: "top", of: document, collision: "fit" },
        open: function (type, data) {

        },
        buttons: [
            {
                text: 'Crear Fase 2',
                click: function () {
                    crearFase2()
                }
            },
            {
                text: 'Cancelar',
                click: function () {
                    $(this).dialog('close')
                }
            }
        ]
    })
}
function sendMailPosVenta()
{
    $.ajax({
        url: 'handlers/Cotizaciones.ashx',
        type: 'POST',
        data: {
            1: 'mailPosVenta',
            to: $('#mail-destinatario').val(),
            asunto: $('#mail-asunto').val(),
            user: localStorage.getItem('user'),
            id: SELECTED_COTIZACION.Id,
            correo: $('#correo2').val()
        },
        success: function (result) {
            if (result.done) {
                alertify.success(result.message)
                $('#send-mail-dialog').dialog('close')
                reloadGridServices(SELECTED_COTIZACION.IT)
                reloadGridCotizaciones();
            } else {
                alertify.error(result.message)
            }
        }
    })
}
function sendMailCotizacion() {
    console.log('mail', 'enviando mail')
    $.ajax({
        url: 'handlers/Cotizaciones.ashx',
        type: 'POST',
        data: {
            1: 'mail-cotizacion',
            asunto: $('#asunto').val(),
            para: $('.para').serializeArray().map(function (m) { return m.value }),
            cc: $('.cc').serializeArray().map(function (m) { return m.value }),
            cco: $('.cco').serializeArray().map(function (m) { return m.value }),
            correo: $('#correo').val(),
            cotizacionId: SELECTED_COTIZACION.Id,
            user: localStorage.getItem('user'),
            // files
            files: FILES.filter(function (f) { return f.checked }).map(function (f) { return f.file }),
        },
        success: function (result) {
            if (result.done) {
                alertify.success(result.message)
                $('#send-mail-cot-dialog').dialog('close')
                reloadGridServices(SELECTED_COTIZACION.IT)
                reloadGridCotizaciones();
            } else {
                alertify.error(result.message)
            }
        }
    })
}
function crearFase2() {
    $.ajax({
        url: 'handlers/Cotizaciones.ashx',
        type: 'POST',
        data: {
            1: 'crearFase2',
            id: SELECTED_COTIZACION.Id,
            date: $('#fecha-fase2').val()
        },
        success: function (result) {
            if (result.done) {
                alertify.success(result.message)
                $('#fase2-dialog').dialog('close')
                reloadGridServices(SELECTED_COTIZACION.IT)
                reloadGridCotizaciones();
                $('#agendar-fase2').hide();
            } else {
                alertify.error(result.message)
            }
        }
    })
}
function aPartirDe() {
    var it = $('#a-partir-de-it').val()
    if (!it) { alertify.error('Ingrese un IT'); return false; }
    $.ajax({
        url: URL_HANDLER,
        data: { 1: 'getOne', it: it },
        success: function (result) {
            if (result.done) {
                setDataForNewCotizacion(result.data)
                $('#add-dialog').dialog('open');
                $('#add-select-dialog').dialog('close')
            } else {
                alertify.error(result.message)
            }
        }
    })
}

function openServiceDetail(rowdata) {
    var it = rowdata.IT;
    reloadGridServices(it);
    $('#services-dialog').dialog('open');
}

function setDataForNewCotizacion(data) {
    SELECTED_CLIENT = data.Cliente
    $('#rut-cliente').val(data.Cliente.Rut);
    $('#nombre-cliente').val(data.Cliente.Nombre);
    $('#giro-cliente').val(data.Cliente.Giro);
    $('#direccion-cliente').val(data.Cliente.Direccion);
    $('#contacto-cliente').val(data.Cliente.NombreContacto);
    $('#email-contacto-cliente').val(data.Cliente.EmailContacto);
    $('#fono-contacto-cliente').val(data.Cliente.TelefonoContacto);
                                // generales
    $('#fecha_doc').val(data.DatosGenerales.FechaDoc);
    $('#fuente_solicitud').val(data.DatosGenerales.FuenteSolicitud);
    $('#forma_pago').val(data.DatosGenerales.FormaDePago);
    $('#moneda').val(data.DatosGenerales.Moneda);
    $('#vendedor').val(data.DatosGenerales.Vendedor);
    $('#observacion').val(data.DatosGenerales.Observacion);
    $('#nota').val(data.DatosGenerales.Nota);
    $('#ciudad').val(data.DatosGenerales.Ciudad);
    $('#validez').val(data.DatosGenerales.FechaValidez);
   // items
    $(data.Items).each(function (i, item) {
        var m = item.ModoDescuento
        var desc = m == '%'
            ? (parseFloat(item.ValorUnitario) * parseFloat(item.Descuento) / 100)
            : m == '$'
                ? parseFloat(item.Descuento) : 0;
        var precioCliente = ((item.ValorUnitario - desc) * item.Cantidad);
        var item1 = {
            Id: Math.random().toString(36).slice(2),
            Producto: item.Producto,
            Descripcion: item.Descripcion,
            Cantidad: item.Cantidad,
            ValorUnitario: item.ValorUnitario,
            ModoDescuento: m,
            Descuento: item.Descuento,
            ValorDescuento: desc,
            PrecioCliente: precioCliente,
            Exento: true,
            TipoElevador: item.TipoElevador,
            TipoElevadorName: item.TipoElevadorName,
            TipoFuncionamientoName: item.TipoFuncionamientoName,
            MarcaName: item.MarcaName,
            TipoFuncionamiento: item.TipoFuncionamiento,
            Marca: item.Marca,
            AnoInstalacion: item.AnoInstalacion,
            EmpresaInstaladora: item.EmpresaInstaladora,
            Ubicacion: item.Ubicacion,
            Uso: item.Uso,
            Altura: item.Altura
        }
        addItem(item1);
    })

    $(data.CostosOperacionales).each(function (c, costo) {
        var total = parseFloat(costo.cantidad) * parseFloat(costo.valor)
        addGasto(costo.cantidad, costo.tipo, costo.tipoName, costo.valor, total)
    })
}

GASTOS = []
function setPriceLogisticCoste(tipoGasto) {
    $.ajax({
        url: URL_HANDLER,
        data: {
            1: 'getPriceLogisticCoste',
            tipoGasto: tipoGasto,
            ciudad: $('#ciudad').val(),
            moneda: $('#moneda').val()
        },
        global: false,
        success: function (result) {
            if (result.done) {
                $('#valor-gasto').val(result.price)
            } else {
                $('#valor-gasto').val(0)
            }
        }
    })
}

function addGasto(cant, tipo, tipoName, valor, total) {
    var id = Math.random().toString(36).slice(2);
    var gasto = { id: id, tipo: tipo, tipoName: tipoName, valor: valor, cantidad: cant, total: total }
    GASTOS.push(gasto)
    $('#gastos-table tbody')
        .append(createViewItemGasto(gasto, false))
    updateValuesLC();
}

function deleteGasto(id, tr) {
    var gasto = Enumerable.from(GASTOS)
        .where(function (w) { return w.id == id })
        .firstOrDefault();
    var index = GASTOS.indexOf(gasto)
    GASTOS.splice(index, 1)
    tr.remove();
    updateValuesLC();
}

function resetAll() {
    $('#form-cliente')[0].reset();
    $('#form-generales')[0].reset();
    $('#form-detalle')[0].reset();
    $('#mail')[0].reset();
    ITEMS = [];
    GASTOS = [];
    $('#grid-items tbody').empty();
    $('.no-validate').removeClass('no-validate');
    $('#grid-items-cantidad').text('')
    $('#grid-items-valor').text('')
    $('#grid-items-descuento').text('')
    $('#grid-items-total').text('')
    $('#total-gastos').text('')
    $('.chk-file').prop('disabled', false);
    $('.chk-file').prop('checked', false);
    $('#no_send_mail').prop('checked', false);
    // Datepicker
    var today = new Date()
    var day = today.getDate()
    var month = today.getMonth() + 1
    var year = today.getFullYear()
    var date = (day < 10 ? '0' + day : day) + '-' + (month < 10 ? '0' + month : month) + '-' + year;
    $('#fecha_doc').val(date)
    getGlobal();
}


function updateValuesLC() {
    var sum = Enumerable.from(GASTOS)
        .sum(function (s) { return s.total })

    $('#total-gastos').text(sum.toLocaleString('es-CL', { maximumFractionDigits: 2, minimumFractionDigits: 0 }));
    $('#total-gastos-resume').text(sum.toLocaleString('es-CL', { maximumFractionDigits: 2, minimumFractionDigits: 0 }));
}
function setPrice() {
    var price = parseFloat($('#item_unitario').val())
    var discountType = $('#item_tipo_descuento').val();
    var discount = $('#item_descuento').val();
    var total
    if (discount != '') {
        if (discountType == 'PORCENTAJE') {
            total = price - (price * parseFloat(discount) / 100);
        } else if (discountType == 'VALOR') {
            total = price - parseFloat(discount)
        } else {
            total = price;
        }

    } else {
        total = price;
    }
    var moneda = parseInt($('#moneda').val())
    $('#item_total').val(total.toFixed(moneda == 1 ? 0 : 2))
}
function getPrice() {
    var producto = $('#producto').val(),
        elevador = $('#tipo_elevador').val(),
        paradas = $('#altura').val(),
        cantidad = $('#item_cantidad').val()
    if (producto == 0 || elevador == 0 || paradas == '' || cantidad == '')
        return;
    $.ajax({
        url: URL_HANDLER,
        global: false,
        data: {
            1: 'setPrice',
            producto: producto,
            elevador: elevador,
            paradas: paradas,
            cantidad: cantidad,
            moneda: $('#moneda').val()
        },
        success: function (result) {
            var price = 0;
            if (result.done) {
                price = result.price;
            }
            $('#item_unitario').val(price);
            setPrice();
        }
    })
}
function setDescripcion() {
    var descripcion = '';
    var producto = $('#producto option:selected').text();
    var tipo_elevador = $('#tipo_elevador option:selected').text();
    var tipo_funcionamiento = $('#tipo_funcionamiento option:selected').text();
    var marca = $('#marca option:selected').text();
    var ano_instalacion = $('#ano_instalacion option:selected').text();
    var empresa_instaladora = $('#empresa_instaladora').val();
    var uso = $('#uso option:selected').text();
    var altura = parseInt($('#altura').val()) || 1;
    var ubicacion = $('#ubicacion').val()
    if (producto) {
        descripcion += producto
    }
    if (tipo_elevador != ' ') {
        descripcion += ' ' + tipo_elevador
    }

    if (tipo_funcionamiento != ' ') {
        descripcion += ' ' + tipo_funcionamiento
    }
    if (uso != ' ') {
        descripcion += ', de uso ' + uso
    }
    if (marca != ' ') {
        descripcion += ', marca ' + marca
    }
    if (ano_instalacion != '') {
        descripcion += ', instalado en ' + ano_instalacion
    }
    if (empresa_instaladora != '') {
        descripcion += ', por la empresa "' + empresa_instaladora + '"'
    }
    if (descripcion != '') {
        descripcion += ' de ' + altura + ' pisos de altura'
    }
    if (ubicacion != '') {
        descripcion += ' ubicado en ' + ubicacion;
    }
    return descripcion.toString().toUpperCase();
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
                $('#nota').val(result.data.nota_descripcion)
                $('#observacion').val(result.data.observacion)
                $('#asunto').val(result.data.mail_envio_cotizacion_subject);
                $('#correo').jqteVal(result.data.mail_envio_cotizacion_text);
                $('#correo2').jqteVal(result.data.mail_post_venta);
            }
        }
    })
}
var FILES = []
function getFiles() {
    $.ajax({
        url: 'handlers/Cotizaciones.ashx',
        data: {
            1: 'get-files',
            type: 'NUEVA_COTIZACION_MAIL'
        },
        success: function (result) {
            if (result.done) {
                var data = result.data
                FILES = [];
                $('#files').empty();
                $.each(data, function (index, file) {
                    var id = Math.random().toString(36).slice(2);
                    FILES.push({ id: id, file: file.Id, checked: false })
                    $('#files')
                        .append($('<li>', { class: 'list-group-item list-group-item-flex' })
                            .append($('<span>', { text: file.Name }))
                            .append($('<div>')
                                .append($('<input>', {
                                    type: 'checkbox',
                                    class: 'chk-file',
                                    checked: false,
                                    change: function () {
                                        var item = Enumerable.from(FILES)
                                            .where(function (w) { return w.id == id })
                                            .firstOrDefault()
                                        item.checked = $(this).is(':checked')
                                    }
                                }))
                                .append($('<button>', {
                                    type: 'button',
                                    class: 'btn btn-danger',
                                    click: function () {
                                        alertify.confirm('Certel', '¿Está seguro de querer deshabilitar este archivo para no ocuparlo en futuras cotizaciones?', function () { disableFile(file.Id) }, function () { })
                                    }
                                }).append($('<i>', { class: 'fa fa-trash' })))
                                )
                        )
                })

            }
        }
    })
}
function disableFile(id) {
    $.ajax({
        url: URL_HANDLER,
        data: { 1: 'disableFile', id: id },
        type: 'POST',
        success: function (result) {
            if (result.done) {
                alertify.success(result.message);
                getFiles()
            } else {
                alertify.error(result.message);
            }
        }
    })
}
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

function setItemsResume() {
    $('#grid-items-resume tbody').empty();
    $.each(ITEMS, function (i, item) {
        $('#grid-items-resume tbody').append(createViewItem(item, true))
    })
    $('#gastos-table-resume tbody').empty();
    $.each(GASTOS, function (i, item) {
        $('#gastos-table-resume tbody').append(createViewItemGasto(item, true))
    })
}
function loadCmbAnoInstalacion() {
    $('#ano_instalacion').empty();
    var year = new Date().getFullYear();
    $('#ano_instalacion').append($('<option>', { val: 0, text: '' }))
    for (var i = year; i > year - 100; i--) {
        $('#ano_instalacion').append($('<option>', { val: i, text: i }))
    }
}
function getPdf(id) {
    $.ajax({
        url: URL_HANDLER,
        data: {
            1: 'get-pdf',
            id: id
        },
        success: function (result) {
            if (result.done) {
                alertify.alert('Certel', 'Documento generado', function () {
                    window.open('/cotizaciones/' + result.filename, '_blank', 'fullscreen=yes')
                })

            }
        }
    })
}
function saveAll() {
    $.ajax({
        url: URL_HANDLER,
        method: 'POST',
        data: {
            1: 'save',
            // cliente
            rut_cliente: $('#rut-cliente').val(),
            nombre_cliente: $('#nombre-cliente').val(),
            giro_cliente: $('#giro-cliente').val(),
            direccion_cliente: $('#direccion-cliente').val(),
            contacto_cliente: $('#contacto-cliente').val(),
            email_contacto_cliente: $('#email-contacto-cliente').val(),
            fono_contacto_cliente: $('#fono-contacto-cliente').val(),
            // generales
            fecha_doc: $('#fecha_doc').val(),
            fuente_solicitud: $('#fuente_solicitud').val(),
            forma_pago: $('#forma_pago').val(),
            moneda: $('#moneda').val(),
            vendedor: $('#vendedor').val(),
            observacion: $('#observacion').val(),
            nota: $('#nota').val(),
            ciudad: $('#ciudad').val(),
            validez: $('#validez').val(),
            // items
            items: JSON.stringify(ITEMS),
            gastos: JSON.stringify(GASTOS),
            // totales

            // mail
            sendMail: !$('#no_send_mail').is(':checked'),
            asunto: $('#asunto').val(),
            para: $('.para').serializeArray().map(function (m) { return m.value }),
            cc: $('.cc').serializeArray().map(function (m) { return m.value }),
            cco: $('.cco').serializeArray().map(function (m) { return m.value }),
            correo: $('#correo').val(),

            // files
            files: FILES.filter(function (f) { return f.checked }).map(function (f) { return f.file }),

            //
            user: localStorage.getItem('user')

        },
        success: function (result) {
            console.log(result)
            if (result.done) {
                alertify.alert('CERTEL', result.message, function () {
                    getPdf(result.id);
                    $('#add-dialog').dialog('close');
                    reloadGridCotizaciones();
                })
            } else {

                alertify.error(result.message);
            }
        }
    })
}
var ITEMS = []
function addItem(item) {
    ITEMS.push(item)
    $('#grid-items tbody').append(createViewItem(item, false))

    updateValues();
    if (CLOSE_DETAIL_DIALOG) {
        $('#form-detalle')[0].reset();
        $('#add-item-dialog').dialog('close');
    }
    alertify.success('Ítem agregado a la cotización')
}
function createViewItem(item, resume) {
    var tr
    var moneda = parseInt($('#moneda').val())
    return tr = $('<tr>')
        .append($('<td>', { class: 'text-right', text: item.Cantidad }))
        .append($('<td>', { text: item.Descripcion }))
        .append($('<td>', { text: item.Ubicacion }))
        .append($('<td>', { class: 'text-right', text: item.ValorUnitario.toLocaleString('es-CL', { maximumFractionDigits: moneda == 1 ? 0 : 2, minimumFractionDigits: 0 }) }))
        .append($('<td>', { text: item.ModoDescuento.toLocaleString('es-CL', { maximumFractionDigits: moneda == 1 ? 0 : 2, minimumFractionDigits: 0 }) }))
        .append($('<td>', { class: 'text-right', text: item.ValorDescuento.toLocaleString('es-CL', { maximumFractionDigits: moneda == 1 ? 0 : 2, minimumFractionDigits: 0 }) }))
        .append($('<td>', { class: 'text-right', text: item.PrecioCliente.toLocaleString('es-CL', { maximumFractionDigits: moneda == 1 ? 0 : 2, minimumFractionDigits: 0 }) }))

        .append(resume ? '' : $('<td>', { class: 'text-center' })
            .append($('<button>', { class: 'btn btn-danger' })
                .append($('<i>', { class: 'fa fa-trash' }))
                .click(function () {
                    console.log('deleting...', item)
                    deleteItemDetail(item.Id, tr)
                })))
}
function createViewItemGasto(item, resume) {
    var tr
    var moneda = parseInt($('#moneda').val())
    return tr = $('<tr>')
        .append($('<td>', { class: 'text-right', text: item.cantidad }))
        .append($('<td>', { text: item.tipoName }))
        .append($('<td>', { class: 'text-right', text: item.valor.toLocaleString('es-CL', { maximumFractionDigits: moneda == 1 ? 0 : 2, minimumFractionDigits: 0 }) }))
        .append($('<td>', { class: 'text-right', text: item.total.toLocaleString('es-CL', { maximumFractionDigits: moneda == 1 ? 0 : 2, minimumFractionDigits: 0 }) }))
        .append(resume ? '' : $('<td>', { class: 'text-center' })
            .append($('<button>', {
                class: 'btn btn-danger',
                click: function () {
                    deleteGasto(item.id, tr);
                }
            }).append($('<i>', { class: 'fa fa-trash' })))
        )
}

function chartQualifyServices(rows) {
    console.log(rows)
    var califica = rows.filter(function (r) { return r.Calificacion == 1 }).length;
    var calificaCO = rows.filter(function (r) { return r.Calificacion == 2 }).length;
    var noCalifica = rows.filter(function (r) { return r.Calificacion == 0 }).length;

    Highcharts.chart('chart', {
        credits: { enabled: false },
        chart: {
            plotBackgroundColor: null,
            plotBorderWidth: null,
            plotShadow: false,
            type: 'pie'
        },
        title: {
            text: 'Calificación de servicios'
        },
        tooltip: {
            pointFormat: '{series.name}: <b>{point.percentage:.1f}%</b>'
        },
        plotOptions: {
            pie: {
                allowPointSelect: true,
                cursor: 'pointer',
                dataLabels: {
                    enabled: true,
                    format: '<b>{point.name}</b>: {point.percentage:.1f} %',
                    style: {
                        color: (Highcharts.theme && Highcharts.theme.contrastTextColor) || 'black'
                    }
                }
            }
        },
        series: [{
            name: 'Servicios',
            colorByPoint: true,
            data: [{
                name: 'Califica',
                y: califica,
                color: '#4CAF50'
            }, {
                name: 'Califica con observaciones',
                y: calificaCO,
                selected: true,
                color: '#f0ad4e'
            }, {
                name: 'No califica',
                y: noCalifica,
                color: '#F44336'
            }]
        }]
    });
}

function updateValues() {
    var cant = Enumerable.from(ITEMS)
        .sum(function (s) { return s.Cantidad });
    var valor = Enumerable.from(ITEMS)
        .sum(function (s) { return s.ValorUnitario });
    var descuento = Enumerable.from(ITEMS)
        .sum(function (s) { return s.ValorDescuento * s.Cantidad });
    var total = Enumerable.from(ITEMS)
        .sum(function (s) { return s.PrecioCliente });
    var recargo = Enumerable.from(GASTOS)
        .sum(function (s) { return s.total })
    var neto = Enumerable.from(ITEMS)
        .sum(function (s) { return (s.ValorUnitario) * s.Cantidad });
    var moneda = parseInt($('#moneda').val())
    $('#grid-items-cantidad').text(cant);
    $('#grid-items-valor').text(neto.toLocaleString('es-CL', { maximumFractionDigits: moneda == 1 ? 0 : 2, minimumFractionDigits: 0 }));
    $('#grid-items-descuento').text(descuento.toLocaleString('es-CL', { maximumFractionDigits: moneda == 1 ? 0 : 2, minimumFractionDigits: 0 }));
    $('#grid-items-total').text(total.toLocaleString('es-CL', { maximumFractionDigits: moneda == 1 ? 0 : 2, minimumFractionDigits: 0 }));

    $('#grid-items-resume-cantidad').text(cant);
    $('#grid-items-resume-valor').text(neto.toLocaleString('es-CL', { maximumFractionDigits: moneda == 1 ? 0 : 2, minimumFractionDigits: 0 }));
    $('#grid-items-resume-descuento').text(descuento.toLocaleString('es-CL', { maximumFractionDigits: moneda == 1 ? 0 : 2, minimumFractionDigits: 0 }));
    $('#grid-items-resume-total').text(total.toLocaleString('es-CL', { maximumFractionDigits: moneda == 1 ? 0 : 2, minimumFractionDigits: 0 }));

    
    var totalTotal = neto - descuento + recargo;
    $('#r_valor').text(neto.toLocaleString('es-CL', { maximumFractionDigits: moneda == 1 ? 0 : 2, minimumFractionDigits: 0 }));
    $('#r_descuento').text(descuento.toLocaleString('es-CL', { maximumFractionDigits: moneda == 1 ? 0 : 2, minimumFractionDigits: 0 }))
    $('#r_recargo').text(recargo.toLocaleString('es-CL', { maximumFractionDigits: moneda == 1 ? 0 : 2, minimumFractionDigits: 0 }))
    $('#r_total').text(totalTotal.toLocaleString('es-CL', { maximumFractionDigits: moneda == 1 ? 0 : 2, minimumFractionDigits: 0 }))

}
function deleteItemDetail(id, div) {
    var obj = Enumerable.from(ITEMS)
        .where(function (w) { return w.Id == id })
        .firstOrDefault();
    var index = ITEMS.indexOf(obj)
    ITEMS.splice(index, 1)
    div.remove();
    updateValues()
}
function setSelectedClient() {
    if (!SELECTED_CLIENT) return;
    $('#rut-cliente').val(SELECTED_CLIENT.Rut)
    $('#direccion-cliente').val(SELECTED_CLIENT.Direccion).focus();
    $('#giro-cliente').val(SELECTED_CLIENT.Giro)
    $('#nombre-cliente').val(SELECTED_CLIENT.Nombre)
    $('#contacto-cliente').val(SELECTED_CLIENT.NombreContacto)
    $('#email-contacto-cliente').val(SELECTED_CLIENT.Email)
    $('#fono-contacto-cliente').val(SELECTED_CLIENT.Telefono)

    // mail
    $('#para').val(SELECTED_CLIENT.Email)
    $('#asunto').val(GLOBAL.mail_envio_cotizacion_subject);
    $('#correo').jqteVal(GLOBAL.mail_envio_cotizacion_text);
}
function getResume() {
    console.log('get resume')
    // cliente
    $('#rc_rut').text($('#rut-cliente').val())
    $('#rc_nombre').text($('#nombre-cliente').val())
    $('#rc_giro').text($('#giro-cliente').val())
    $('#rc_direccion').text($('#direccion-cliente').val())
    $('#rc_contacto').text($('#contacto-cliente').val())
    $('#rc_email').text($('#email-contacto-cliente').val())
    $('#rc_fono').text($('#fono-contacto-cliente').val())
    // generales
    $('#rg_fecha_doc').text($('#fecha_doc').val())
    $('#rg_fuente_solicitud').text($('#fuente_solicitud option:selected').text())
    $('#rg_forma_de_pago').text($('#forma_pago option:selected').text())
    $('#rg_moneda').text($('#moneda option:selected').text())
    $('#rg_vendedor').text($('#vendedor option:selected').text())
    $('#rg_observaciones').text($('#observacion').val())
    $('#rg_nota').text($('#nota').val())
    $('#rg_validez').text($('#validez').val())
    $('#rc_ciudad').text($('#ciudad option:selected').text())
    updateValues();
    setItemsResume();

}
function reloadGridClientes() {
    $("#grid-clientes").jqGrid('setGridParam', {
        url: URL_HANDLER,
        postData: {
            1: 'grid-clientes',
            name: $('#fc_nombre').val(),
            rut: $('#fc_rut').val()
        }
    }).trigger('reloadGrid');
}

function reloadGridServices(it) {
    $("#services-grid").jqGrid('setGridParam', {
        url: URL_HANDLER,
        postData: {
            1: 'getServicesFromIT',
            it: it
        }
    }).trigger('reloadGrid');
}
function reloadGridProductos() {
    $("#grid-productos").jqGrid('setGridParam', {
        url: URL_HANDLER,
        postData: {
            1: 'grid-productos',
            name: $('#fp_nombre').val()
        }
    }).trigger('reloadGrid');
}
function reloadGridCotizaciones() {
    $("#grid").jqGrid('setGridParam', {
        url: URL_HANDLER,
        page: 1,
        postData: {
            1: 'grid-cotizaciones',
            cliente: $('#f_name').val(),
            estado: $('#f_estado').val(),
            inicio: $('#f_inicio').val(),
            fin: $('#f_fin').val(),
            alertF2: $('#f_alertF2').is(':checked')
        }
    }).trigger('reloadGrid');
}
function getClientByRut() {
    if (!$('#rut-cliente').val()) {
        return;
    }
    $.ajax({
        url: URL_HANDLER,
        data: { 1: 'getClientByRut', rut: $('#rut-cliente').val() },
        success: function (result) {
            if (result.done) {
                if (result.client) {
                    SELECTED_CLIENT = result.client;
                    setSelectedClient()
                }
            }
        }
    })
}