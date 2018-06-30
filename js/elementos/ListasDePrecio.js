
$(document).ready(function () {

    combobox($('#n_tipo-elevador'), { 1: 'aparatos' }, 'Seleccione...');
    combobox($('#m_tipo-elevador'), { 1: 'aparatos' }, 'Seleccione...');

    combobox($('#n_tipo-Producto'), { 1: 'productos' }, 'Seleccione...');
    combobox($('#m_tipo-Producto'), { 1: 'productos' }, 'Seleccione...');


    combobox($('#f_producto'), { 1: 'productos' }, 'Todos...');
    combobox($('#f_elevador'), { 1: 'aparatos' }, 'Todos...');

    $('#f_remove').click(function () {
        $('#filter-form')[0].reset();
        reloadGrid();
    })
    $('#filter-form').submit(function () {
        reloadGrid();
        return false;
    })
    $("#grid").jqGrid({
        url: 'handlers/ListasDePrecio.ashx',
        postData: {
            1: 'grid',
            product: 0,
            paradas: 0,
            moreThanOne: true,
            elevador: 0
        },
        mtype: "POST",
        autowidth: true,
        datatype: "json",
        shrinkToFix: true,
        ajaxGridOptions: { global: false },
        colModel: [
            { label: 'Id', name: 'Id', key: true, hidden: true },
            { label: 'Elevador', name: 'Aparato', index: 'Aparato.Nombre', width: 10, align: 'left', sortable: true },
            { label: 'Paradas', name: 'Paradas', index: 'Paradas', width: 5, align: 'center', sortable: true },
            { label: 'Mas De 1 Equipo', name: 'MasDeUnEquipo', index: 'MasDe1Equipo', width: 5, align: 'center', sortable: true },
            { label: 'Producto', name: 'TipoProducto', index: 'Producto.Descripcion', width: 10, align: 'left', sortable: true },
            { label: 'Valor UF', name: 'valor', index: 'ValorUF', width: 10, align: 'right', sortable: true },
            { name: 'Eliminar', width: 5, align: 'center', editable: false, sortable: false, formatter: btnEliminar },
            { name: 'Modificar', width: 5, align: 'center', editable: false, sortable: false, formatter: btnModificar },
            { label: '', name: 'AparatoID', index: 'AparatoID', hidden: true },
            { label: '', name: 'TipoProductoID', index: 'TipoProductoID', hidden: true },
        ],
        sortname: 'Id', sortorder: 'desc',
        viewrecords: true,
        //autoheight: true,
        height: 400,
        rowNum: 30,
        pager: "#pager",
        caption: '',

        onCellSelect: function (rowid, iCol, cellcontent, e) {
            var rowData = $(this).jqGrid('getRowData', rowid);
            SELECTED = rowData;

            var ID = $("#grid").getCell(rowid, 'Id');
            var AparatoID = $("#grid").getCell(rowid, 'AparatoID');
            var TipoProductoID = $("#grid").getCell(rowid, 'TipoProductoID');
            var valor = $("#grid").getCell(rowid, 'valor');
            var MasDeUnEquipo = $("#grid").getCell(rowid, 'MasDeUnEquipo');
            var Paradas = $("#grid").getCell(rowid, 'Paradas');
            if (iCol == 6) {
                alertify.confirm('Eliminar', '¿Esta seguro que desea eliminar la lista de precio seleccionada?',
                    function () {

                        $.ajax({
                            url: 'handlers/ListasDePrecio.ashx',
                            data: {
                                1: 'delPriceList',
                                id: ID,
                            },
                            success: function (result) {
                                if (result.done) {
                                    alertify.success(result.message);
                                    reloadGrid();
                                }
                                else
                                    alertify.error(result.message);
                            }
                        });

                    },
                    function () { });


            } else if (iCol == 7) {
                ModificarLista(ID, TipoProductoID, Paradas, MasDeUnEquipo, AparatoID, valor);
            }

        }
    }).navGrid('#pager',
        {
            edit: false, add: false, del: false, search: false,
            refresh: true, view: false, position: "left", cloneToTop: false
        });

    function btnModificar(cellvalue, options, rowObject) {
        return '<div style="cursor: pointer"><i class="fa fa-pencil fa-2x"></i></div>';
    };

    function btnEliminar(cellvalue, options, rowObject) {
        return div = '<div style="cursor: pointer"><i class="fa fa-trash fa-2x" style="color: #E61414"></i></div>';
    };

    $('#add-lista').click(function () {

        $('#dialog-nueva-lista').dialog({
            modal: true, bgiframe: false, width: 600, title: 'Nueva Lista de precio', draggable: true,
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
                        $('#form-nueva-lista').submit();
                    }
                }
            ]
        });
    });

    $('#form-nueva-lista').submit(function (e) {
        e.preventDefault();
        if (!validateForm($(this))) {
            alertify.error('Complete los datos requeridos');
            return false;
        }

        $.ajax({
            url: 'handlers/ListasDePrecio.ashx',
            data: {
                1: 'addPriceList',
                TipoElevadorId: $('#n_tipo-elevador').val(),
                Paradas: $('#n_Paradas').val(),
                MasDe1Equipo: $('#n_mas-de-un-equipo').val(),
                ProductoId: $('#n_tipo-Producto').val(),
                ValorUF: $('#n_valor').val()
            },
            success: function (result) {
                if (result.done) {
                    alertify.success(result.message);
                    $('#dialog-nueva-lista').dialog('close');
                    reloadGrid();
                }
                else
                    alertify.error(result.message);
            }
        });
        return false;

    });

    function ModificarLista(ID, TipoProductoID, Paradas, MasDeUnEquipo, AparatoID, valor) {
        $('#dialog-mod-lista').dialog({
            modal: true, bgiframe: false, width: 600, title: 'Modificar Lista de precio', draggable: true,
            resizable: false, closeOnEscape: true, autoOpen: true, show: 'clip', hide: 'puff',
            open: function (type, data) {
                $(this).find('form')[0].reset();

                $('#m_id').val(ID)
                $('#m_valor').val(valor)
                $('#m_Paradas').val(Paradas)
                $("#m_tipo-Producto").val(TipoProductoID);
                $("#m_tipo-elevador").val(AparatoID);
                if (MasDeUnEquipo == 'Si')
                    $("#m_mas-de-un-equipo").prop('checked', true);
                else
                    $("#m_mas-de-un-equipo").prop('checked', false);
            },
            close: function () {

            },
            buttons: [
                {
                    id: 'btn-add',
                    text: 'Guardar',
                    click: function () {
                        $('#form-mod-lista').submit();
                    }
                }
            ]
        });
    }

    $('#form-mod-lista').submit(function (e) {
        e.preventDefault();
        if (!validateForm($(this))) {
            alertify.error('Complete los datos requeridos');
            return false;
        }

        $.ajax({
            url: 'handlers/ListasDePrecio.ashx',
            data: {
                1: 'modPriceList',
                id: $('#m_id').val(),
                TipoElevadorId: $('#m_tipo-elevador').val(),
                Paradas: $('#m_Paradas').val(),
                MasDe1Equipo: $('#m_mas-de-un-equipo').val(),
                ProductoId: $('#m_tipo-Producto').val(),
                ValorUF: $('#m_valor').val()
            },
            success: function (result) {
                if (result.done) {
                    alertify.success(result.message);
                    $('#dialog-mod-lista').dialog('close');
                    reloadGrid();
                }
                else
                    alertify.error(result.message);
            }
        });
        return false;

    });
});

function reloadGrid() {
    jQuery("#grid")
        .jqGrid('setGridParam',
        {
            url: 'handlers/ListasDePrecio.ashx',
            postData: {
                1: 'grid',
                product: $('#f_producto').val(),
                paradas: $('#f_paradas').val(),
                moreThanOne: $('#f_moreThanOne').is(':checked'),
                elevador: $('#f_elevador').val()
            },
        })
        .trigger("reloadGrid");
}
