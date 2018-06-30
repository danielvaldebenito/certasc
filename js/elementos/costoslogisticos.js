var URL_HANDLER = 'handlers/CostosLogisticos.ashx'
var SELECTED_ID
$(document).ready(function () {
    // combos
    combobox($('#f_ciudad'), { 1: 'ciudades' }, 'Todas...');
    combobox($('#add-ciudad'), { 1: 'ciudades' }, 'Seleccione...');
    combobox($('#edit-ciudad'), { 1: 'ciudades' }, '');
    combobox($('#f_tipoCosto'), { 1: 'tipoGasto' }, 'Todos...');
    combobox($('#add-tipo-costo'), { 1: 'tipoGasto' }, 'Seleccione...');
    combobox($('#edit-tipo-costo'), { 1: 'tipoGasto' }, '');


    // events
    $('#add-type').click(function () {
        $('#add-type-dialog').dialog('open')
    })

    // grid
    $('#grid').jqGrid({
        url: URL_HANDLER,
        postData: {
            1: 'grid-costos-logisticos',
            ciudad: $('#f_ciudad').val() || 0,
            tipoCosto: $('#f_tipoCosto').val() || 0
        },
        mtype: "POST",
        autowidth: true,
        datatype: "json",
        shrinkToFix: true,
        ajaxGridOptions: { global: false },
        colModel: [
            { label: 'Id', name: 'Id', index: 'Id', key: true, hidden: true, width: 10, align: 'center', sortable: true },
            { label: 'Ciudad', name: 'Ciudad', index: 'Ciudad.Descripcion', width: 20, align: 'center', sortable: true },
            { label: 'Tipo Costo', name: 'TipoGasto', index: 'TipoGastoId', sortable: true, width: 20, align: 'left' },
            { label: 'Valor', name: 'Valor', index: 'ValorCLP', width: 20, align: 'right', sortable: true, formatter: 'number', formatoptions: { decimalSeparator: ",", thousandsSeparator: ".", decimalPlaces: 0 } },
            { label: '', name: '', width: 10, align: 'center', formatter: btnEdit },
            { label: '', name: '', width: 10, align: 'center', formatter: btnDelete },
            { name: 'CiudadId', hidden: true },
            { name: 'TipoCostoId', hidden: true }
        ],
        sortname: 'Ciudad.Descripcion',
        sortorder: 'asc',
        viewrecords: true,
        height: 400,
        rowNum: 30,
        pager: "#pager",
        caption: 'Lista de costos logísticos',
        onCellSelect: function (rowid, iCol, cellcontent, e) {
            var rowData = $(this).jqGrid('getRowData', rowid);
            if (iCol == 4) {
                
                $('#edit-ciudad').val(rowData.CiudadId);
                $('#edit-tipo-costo').val(rowData.TipoCostoId);
                $('#edit-value').val(rowData.Valor);
                SELECTED_ID = rowData.Id
                $('#edit-dialog').dialog('open')
            }
            if (iCol == 5) {
                alertify.confirm('Certel', '¿Está seguro de eliminar este ítem?', function () { deleteOne(rowData.Id) }, function () { })
            }
        }
    }).navGrid('#pager',
        {
            edit: false, add: false, del: false, search: false,
            refresh: true, view: false, position: "left", cloneToTop: false
        });

    $('#filter-form').submit(function () {
        reloadGrid();
        return false;
    })
    $('#filter-form select').change(function () {
        $('#filter-form').submit();
    })
    $('#f_remove').click(function () {
        $('#filter-form')[0].reset();
        $('#filter-form').submit();
    })
    $('#add-dialog').dialog({
        modal: true, bgiframe: false, width: 400, title: 'Nuevo ítem', draggable: true,
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
    $('#add-type-dialog').dialog({
        modal: true, bgiframe: false, width: 400, title: 'Nuevo tipo', draggable: true,
        resizable: false, closeOnEscape: true, autoOpen: false, show: 'clip', hide: 'puff',
        position: { my: "center", at: "top", of: document, collision: "fit" },
        open: function (type, data) {

        },
        close: function () {
            $('#add-type-form')[0].reset();
        },
        buttons: [
            {
                text: 'Guardar',
                click: function () {
                    $('#add-type-form').submit()
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
        modal: true, bgiframe: false, width: 400, title: 'Editar ítem', draggable: true,
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
    $('#add-form').submit(function () {
        if (!validateForm($(this))) {
            alertify.error('Complete los datos requeridos');
            return false;
        }
        add();
        return false;
    })
    $('#add-type-form').submit(function () {
        if (!validateForm($(this))) {
            alertify.error('Complete los datos requeridos');
            return false;
        }
        addType();
        return false;
    })
    $('#edit-form').submit(function () {
        if (!validateForm($(this))) {
            alertify.error('Complete los datos requeridos');
            return false;
        }
        edit(SELECTED_ID)
        return false;
    })
    function btnEdit() {
        return '<div class="btn-grid"><i class="fa fa-pencil"></i></div>'
    }
    function btnDelete() {
        return '<div class="btn-grid"><i class="fa fa-trash" style="color: red"></i></div>'
    }
    $('#add').click(function () {
        $('#add-dialog').dialog('open');
    })
})

function reloadGrid() {
    $('#grid').jqGrid('setGridParam', {
        url: URL_HANDLER,
        postData: {
            1: 'grid-costos-logisticos',
            ciudad: $('#f_ciudad').val() || 0,
            tipoCosto: $('#f_tipoCosto').val() || 0
        },
    }).trigger('reloadGrid')
}
function addType()
{
    $.ajax({
        url: URL_HANDLER,
        type: 'POST',
        data: {
            1: 'addType', name: $('#add-type-name').val()
        },
        success: function (result) {
            if (result.done) {
                alertify.success(result.message);
                combobox($('#add-tipo-costo'), { 1: 'tipoGasto' }, 'Seleccione...');
                combobox($('#f_tipoCosto'), { 1: 'tipoGasto' }, 'Todos...');
                combobox($('#edit-tipo-costo'), { 1: 'tipoGasto' }, 'Seleccione...');
                $('#add-type-dialog').dialog('close')
                setTimeout(function () { $('#add-tipo-costo').val(result.id) }, 1000)

            } else {
                alertify.success(result.message);
            }
        }
    })
}
function add() {
    $.ajax({
        url: URL_HANDLER,
        type: 'POST',
        data: { 1: 'add', ciudad: $('#add-ciudad').val(), tipoGasto: $('#add-tipo-costo').val(), valor: $('#add-value').val() },
        success: function (result) {
            if (result.done) {
                alertify.success(result.message);
                reloadGrid();
                $('#add-dialog').dialog('close')
            } else {
                alertify.success(result.message);
            }
        }
    })
}
function edit(id) {
    $.ajax({
        url: URL_HANDLER,
        type: 'POST',
        data: { 1: 'edit', id: id, valor: $('#edit-value').val() },
        success: function (result) {
            if (result.done) {
                alertify.success(result.message);
                reloadGrid();
                $('#edit-dialog').dialog('close')
            } else {
                alertify.success(result.message);
            }
        }
    })
}
function deleteOne (id){
    $.ajax({
        url: URL_HANDLER,
        type: 'POST',
        data: { 1: 'delete', id: id },
        success: function (result) {
            if (result.done) {
                alertify.success(result.message);
                reloadGrid();
                
            } else {
                alertify.success(result.message);
            }
        }
    })
}
