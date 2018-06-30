var SELECTED;
$(document).ready(function () {

    $('#ac_rut, #ec_rut, #f_rut').Rut({
        format_on: 'keyup'
    })
    $("#grid").jqGrid({
        url: 'handlers/Clientes.ashx',
        postData: {
            1: 'grid',
            name: $('#f_name').val(),
            rut: $('#f_rut').val()
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
            { label: 'Giro', name: 'Giro', index: 'Giro', width: 30, align: 'left', sortable: true },
            { label: 'Dirección', name: 'Direccion', index: 'Direccion', width: 35, align: 'left', sortable: true },
            { label: 'Teléfono', name: 'Telefono', index: 'Telefono', width: 30, align: 'left', sortable: true },
            { label: 'E-mail', name: 'Email', index: 'Email', width: 30, align: 'left', sortable: true },
            { label: 'Editar', width: 10, formatter: editButton, align: 'center' },
            { label: 'Habilitado', name: 'Enabled', width: 10, formatter: enableButton, align: 'center' },
            { label: 'Cotizaciones', name: '', width: 10, formatter: cotButton, align: 'center' }
        ],
        sortname: 'FechaCreacion', sortorder: 'desc',
        viewrecords: true,
        height: 350,
        rowNum: 30,
        pager: "#pager",
        caption: 'Clientes registrados en el sistema',
        
        onCellSelect: function (rowid, iCol, cellcontent, e) {
            var rowData = $(this).jqGrid('getRowData', rowid);
            SELECTED = rowData;
            if(iCol == 7) {
                openDialogEdit(rowData);
            }
            if(iCol == 8) {
                enableOrDisabled(rowData);
            }
            if (iCol == 9) {
                window.location.href = 'handlers/ExportExcel.ashx?1=cotizacionesCliente&id=' + rowData.Id
            }
        }
    }).navGrid('#pager',
                {
                    edit: false, add: false, del: false, search: false,
                    refresh: true, view: false, position: "left", cloneToTop: false
                });
    $("#grid").navButtonAdd('#pager', {
        caption: "Exportar",
        title: "Exportar base de datos de clientes a Excel",
        buttonicon: "ui-icon-document",
        onClickButton: function () {
            window.location.href = 'handlers/ExportExcel.ashx?1=clientes&name=' + $('#f_name').val() + '&rut=' + $('#f_rut').val()
        },
        position: "first"
    });
    $('#formFiltros').submit(function (e) {
        e.preventDefault();
        jQuery("#grid")
            .setGridParam({
                postData: {
                    1: 'grid',
                    name: $('#f_name').val(),
                    rut: $('#f_rut').val()
                },
            })
            .trigger("reloadGrid");
        return false;
    })
    .submit();
    $('#f_remove').click(function () {
        $('#formFiltros')[0].reset();
        $('#formFiltros').submit();
    });

    // Button
    function editButton ()
    {
        return '<div class="btn-grid"><i class="fa fa-pencil"></i></div>'
    }
    function enableButton(cellvalue, options, rowObject) {
        var color = cellvalue == true ? 'green' : 'red';
        return '<div class="btn-grid"><i class="fa fa-circle" style="color:' + color + '"></i></div>'
    }
    function cotButton(cellvalue, options, rowObject) {
        return '<div class="btn-grid"><i class="fa fa-file-excel-o" style="color:green"></i></div>'
    }
    $('#add-client-form').submit(function (e) {
        e.preventDefault();
        var validateRut = $.Rut.validar($('#ac_rut').val());
        if (!validateRut)
        {
            alertify.error('Rut Ingresado no es válido');
            return;
        }
        if (!validateForm($(this))) {
            alertify.error('Complete los datos requeridos');
            return;
        }
        $.ajax({
            url: 'handlers/Servicios.ashx',
            data: {
                1: 'addClient',
                rut: $('#ac_rut').val(),
                nombre: $('#ac_nombre').val(),
                giro: $('#ac_giro').val(),
                direccion: $('#ac_direccion').val(),
                telefono: $('#ac_telefono').val(),
                email: $('#ac_mail').val()
            },
            success: function (result) {
                if (result.done) {
                    alertify.success(result.message);
                    $('#add-client-dialog').dialog('close');
                    $('#formFiltros').submit();
                }
                else
                    alertify.error(result.message);
            }
        });
        return false;
    });

    // Edit
    $('#edit-client-form').submit(function (e) {
        e.preventDefault();

        if (!validateForm($(this))) {
            alertify.error('Complete los datos requeridos');
            return;
        }
        var validateRut = $.Rut.validar($('#ec_rut').val());
        if (!validateRut) {
            alertify.error('Rut Ingresado no es válido');
            return;
        }
        $.ajax({
            url: 'handlers/Clientes.ashx',
            data: {
                1: 'editClient',
                rut: $('#ec_rut').val(),
                nombre: $('#ec_nombre').val(),
                giro: $('#ec_giro').val(),
                direccion: $('#ec_direccion').val(),
                telefono: $('#ec_telefono').val(),
                email: $('#ec_mail').val(),
                id: SELECTED.Id
            },
            success: function (result) {
                if (result.done) {
                    alertify.success(result.message);
                    $('#edit-client-dialog').dialog('close');
                    $('#formFiltros').submit();
                }
                else
                    alertify.error(result.message);
            }
        });
        return false;
    });

    $('#add-client').click(function () {
        $('#add-client-dialog').dialog({
            modal: true, bgiframe: false, width: '50%', title: 'Nuevo Cliente', draggable: true,
            resizable: false, closeOnEscape: true, autoOpen: true, show: 'clip', hide: 'puff',
            open: function (type, data) {

                $(this).find('form')[0].reset();
                $('.no-validate').removeClass('no-validate')
            },
            close: function () {

            },
            buttons: [{
                id: 'newClientSave',
                text: 'Guardar',
                click: function () {
                    $('#add-client-form').submit();
                }
            }, {
                id: 'newClientClose',
                text: 'Cerrar',
                click: function () {
                    $('#add-client-dialog').dialog('close');
                }
            }]
        });
    });

});

function openDialogEdit(rowData)
{
    $('#ec_nombre').val(rowData.Nombre);
    $('#ec_rut').val(rowData.Rut);
    $('#ec_telefono').val(rowData.Telefono);
    $('#ec_direccion').val(rowData.Direccion);
    $('#ec_mail').val(rowData.Email);
    $('#ec_giro').val(rowData.Giro);

    $('#edit-client-dialog').dialog({
        modal: true, bgiframe: false, width: '50%', title: 'Editar Cliente', draggable: true,
        resizable: false, closeOnEscape: true, autoOpen: true, show: 'clip', hide: 'puff',
        open: function (type, data) {

        },
        close: function () {

        },
        buttons: [{
            id: 'editClientSave',
            text: 'Guardar',
            click: function () {
                $('#edit-client-form').submit();
            }
        }, {
            id: 'editClientClose',
            text: 'Cerrar',
            click: function () {
                $('#edit-client-dialog').dialog('close');
            }
        }]
    });
}

function enableOrDisabled(rowData)
{
    $.ajax({
        url: 'handlers/Clientes.ashx',
        type: 'POST',
        data: { 1: 'enabledOrDisabled', id: rowData.Id },
        success: function (result) {
            if (result.done)
            {
                $('#formFiltros').submit();
                alertify.success(result.message);
            }
            else
            {
                alertify.error(result.message);
            }
            
        }
    })
}