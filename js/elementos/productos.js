var URL_HANDLER = 'handlers/productos.ashx';
var SELECTED_PRODUCT = {}
if (!localStorage.getItem('user')) {
    window.location.href = '/'
}
$(document).ready(function () {

    // events
    $('#f_name').keyup(function () {
        $('#filter-form').submit();
    })
    $('#add').click(function () {
        $('#add-dialog').dialog('open')
    })
    // dialogs
    $('#add-dialog').dialog({
        modal: true, bgiframe: false, width: 400, title: 'Nuevo producto', draggable: true,
        resizable: false, closeOnEscape: true, autoOpen: false, show: 'clip', hide: 'puff',
        position: { my: "center", at: "top", of: document, collision: "fit" },
        open: function (type, data) {

        },
        close: function () {
            $('#add-form')[0].reset();
        },
        buttons: [
            {
                text: 'Guardar',
                click: function () {
                    $('#add-form').submit()
                }
            },
            {
                text: 'Cancelar',
                click: function () {
                    $(this).dialog('close')
                }
            }]
    });
    $('#edit-dialog').dialog({
        modal: true, bgiframe: false, width: 400, title: 'Editar producto', draggable: true,
        resizable: false, closeOnEscape: true, autoOpen: false, show: 'clip', hide: 'puff',
        position: { my: "center", at: "top", of: document, collision: "fit" },
        open: function (type, data) {

        },
        close: function () {
            $('#edit-form')[0].reset();
        },
        buttons: [
            {
                text: 'Guardar',
                click: function () {
                    $('#edit-form').submit()
                }
            },
            {
                text: 'Cancelar',
                click: function () {
                    $(this).dialog('close')
                }
            }]
    });

    // forms
    $('#filter-form').submit(function () {
        reloadGrid();
        return false;
    })
    $('#add-form').submit(function () {
        if (!validateForm($(this))) {
            alertify.error('Complete los datos requeridos')
            return false;
        }
        add();
        return false;
    })
    $('#edit-form').submit(function () {
        if (!validateForm($(this))) {
            alertify.error('Complete los datos requeridos')
            return false;
        }
        edit();
        return false;
    })
    // grid
    $('#grid').jqGrid({
        url: URL_HANDLER,
        postData: {
            1: 'grid',
            name: $('#f_name').val()
        },
        mtype: "POST",
        autowidth: true,
        datatype: "json",
        shrinkToFix: true,
        ajaxGridOptions: { global: false },
        colModel: [
            { label: 'Id', name: 'Id', index: 'Id', key: true, hidden: true, width: 10, align: 'center', sortable: true },
            { label: 'Nombre', name: 'Nombre', index: 'Descripcion', width: 200, align: 'center', sortable: true},
            { label: 'Editar', formatter: btnEdit, width: 10, align: 'center' },
            { label: 'ON/OFF', name: 'Habilitado', formatter: btnRemove, width: 10, align: 'center' }
        ],
        sortname: 'Id',
        sortorder: 'asc',
        viewrecords: true,
        height: 400,
        rowNum: 30,
        pager: "#pager",
        caption: 'Lista de productos',
        onCellSelect: function (rowid, iCol, cellcontent, e) {
            var rowData = $(this).jqGrid('getRowData', rowid);
            SELECTED_PRODUCT = rowData;
            if (iCol == 2) {
                $('#edit-name').val(rowData.Nombre)
                $('#edit-dialog').dialog('open')
            }
            if (iCol == 3) {
                remove(rowData.Id)
            }
        }
    }).navGrid('#pager',
        {
            edit: false, add: false, del: false, search: false,
            refresh: true, view: false, position: "left", cloneToTop: false
        });

    function btnEdit() {
        return '<div class="btn-grid"><i class="fa fa-pencil"></i></div>'
    }
    function btnRemove(cellvalue, options, rowObject) {
        var color = cellvalue ? 'green' : 'red'
        return '<div class="btn-grid"><i class="fa fa-circle" style="color: ' + color + '"></i></div>'
    }
})

function add() {
    $.ajax({
        url: URL_HANDLER,
        type: 'POST',
        data: {
            1: 'add',
            name: $('#add-name').val()
        },
        success: function (result) {
            if (result.done) {
                alertify.success(result.message);
                reloadGrid();
                $('#add-dialog').dialog('close');
            } else {
                alertify.error(result.message);
            }
        }
    })
}
function edit() {
    $.ajax({
        url: URL_HANDLER,
        type: 'POST',
        data: {
            1: 'edit',
            name: $('#edit-name').val(),
            id: SELECTED_PRODUCT.Id
        },
        success: function (result) {
            if (result.done) {
                alertify.success(result.message);
                reloadGrid();
                $('#edit-dialog').dialog('close');
            } else {
                alertify.error(result.message);
            }
        }
    })
}
function remove() {
    $.ajax({
        url: URL_HANDLER,
        type: 'POST',
        data: {
            1: 'enableOrDisable',
            id: SELECTED_PRODUCT.Id
        },
        success: function (result) {
            if (result.done) {
                alertify.success(result.message);
                reloadGrid();
            } else {
                alertify.error(result.message);
            }
        }
    })
}
function reloadGrid() {
    $('#grid').jqGrid('setGridParam', {
        url: URL_HANDLER,
        postData: {
            1: 'grid',
            name: $('#f_name').val()
        },
    }).trigger('reloadGrid')
}