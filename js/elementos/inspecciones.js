var CLOSE_DIALOG;
var SELECTED_INSPECCION;
var SELECTED_CARACTERISTICA;
var SPECIFIC_DATA_APPLY_ALL = false;
var SELECTED_INSPECCION_ROW = {};
var SELECTED_OT = {};
$(document).ready(function () {

    $('[name="check"]').bootstrapSwitch();
    $('[name="copy-inspection"]').bootstrapSwitch({
        onText: 'SÍ',
        offText: 'NO',
        onSwitchChange: function (event, state) {
            var value = parseInt(event.currentTarget.value);
            var toChange = state ? value - 1 : value + 1;
            $('[name="copy-inspection"][value="' + toChange + '"]').bootstrapSwitch('state', state);
            
        }
    });
    $('[name="copy-new-inspection"]').bootstrapSwitch({
        onText: 'Nueva',
        offText: 'Existente',
        state: true,
        onSwitchChange: function (event, state) {
            if(state)
            {
                $('#copy-to-it-new').show();
                $('#copy-to-it-exist').hide();
            }
            else {
                $('#copy-to-it-new').hide();
                $('#copy-to-it-exist').show();
            }
        }
    });
    
    // datepickers
    $('#f_desde, #f_hasta, #ai-fecha-inspeccion').datepicker();
    $('#ei-fecha-inspeccion').datepicker();
    var currrentYear = new Date().getFullYear()
    $('#ai-fecha-instalacion, #ei-fecha-instalacion, #ei-fec, #ai-fec, #ei-fvc, #ai-fvc').datepicker({ changeYear: true, yearRange: "1935:2030" });
    $('#fecha_entrega').datepicker({ formatDate: 'dd-mm-yy', minDate: 0 });
    // comboboxs
    combobox($('#ai-aparato'), { 1: 'aparatos' }, 'Seleccione...');
    combobox($('#ai-ingeniero'), { 1: 'inspectores' }, 'Seleccione...');
    combobox($('#ai-tipo-funcionamiento'), { 1: 'tipoFuncionamiento' }, 'Seleccione...');
    combobox($('#ai-destino'), { 1: 'destinoProyecto' }, 'Seleccione...');

    combobox($('#ei-aparato'), { 1: 'aparatos' }, '');
    combobox($('#ei-ingeniero'), { 1: 'inspectores' }, '');
    combobox($('#ei-tipo-funcionamiento'), { 1: 'tipoFuncionamiento' }, '');
    combobox($('#ei-destino'), { 1: 'destinoProyecto' }, 'Seleccione...');
    
    $('#crearfase2').bootstrapSwitch({
        onText: 'SÍ',
        offText: 'NO',
        state: false,
        onSwitchChange: function (event, state) {
            
        }
    });
    $('input[name=calificacion]').change(function () {
        var value = $(this).val();
        console.log('change', value);
        if (value == 1)
        {
            $('#plazo').hide();
            $('#fase2').hide();
        }   
        else
        {
            $('#plazo').show();
            $('#fase2').show(); 
        }
            
    });
    $("#grid").jqGrid({
        url: 'handlers/Inspecciones.ashx',
        postData: {
            1: 'grid',
            it: $('#f_it').val(),
            desde: $('#f_desde').val(),
            hasta: $('#f_hasta').val(),
            calificacion: $('#f_calificacion').val(),
            roles: localStorage.getItem('roles'),
            user: localStorage.getItem('user')
        },
        mtype: "POST",
        autowidth: true,
        datatype: "json",
        shrinkToFix: true,
        ajaxGridOptions: { global: false },
        colModel: [
            { label: 'Id', name: 'Id', key: true, hidden: true },
            { label: 'IT', name: 'It', index: 'IT', width: 25, align: 'center', sortable: false },
            { label: 'Elevador', name: 'Aparato', index: 'Aparato.Nombre', width: 30, align: 'center', sortable: true },
            { label: 'Funcionamiento', name: 'Funcionamiento', index: 'TipoFuncionamientoAparato.Descripcion', width: 35, align: 'center', sortable: true },
            { label: 'F. Creación', name: 'FechaCreacion', index: 'FechaCreacion', width: 25, align: 'center', sortable: true },
            { label: 'F. Inspección', name: 'FechaInspeccion', index: 'FechaInspeccion', width: 25, align: 'center', sortable: true },
            { label: 'Fase', name: 'Fase', index: 'Fase', width: 10, align: 'center', sortable: true, formatter: fase },
            { label: 'Norma', name: 'Norma', index: 'Norma', width: 30, align: 'center', sortable: true },
            { label: 'Ingeniero', name: 'Ingeniero', index: 'Ingeniero', width: 30, align: 'center', sortable: true },
            //{ label: 'Estado', name: 'Estado', index: 'EstadoInspeccion.Descripcion', width: 15, formatter: btnEstado, align: 'center', sortable: true, hidden: true },
            { label: 'Editar', name: '', width: 15, formatter: btnEditar, align: 'center', sortable: false },
            { label: 'C. Particulares', name: '', width: 15, formatter: btnEspecifico, align: 'center', sortable: false },
            { label: 'Normas', name: '', width: 15, formatter: btnNormas, align: 'center', sortable: false },
            { label: 'Check-List', name: '', width: 15, formatter: btnCheckList, align: 'center', sortable: false },
            { label: 'Obs-Tec', name: 'ObsTec', width: 15, formatter: btnObsTec, align: 'center', sortable: false },
            { label: 'Califica', name: '', width: 15, formatter: btnCalifica, align: 'center', sortable: false },
            { label: '', name: 'Revisar', width: 15, formatter: btnRevisado, align: 'center', sortable: false },
            { label: '', name: 'Aprobar', width: 15, formatter: btnAprobado, align: 'center', sortable: false },
            { label: 'Informe', name: '', width: 15, formatter: btnInforme, align: 'center', sortable: false },
            { label: 'Fin', name: '', width: 15, formatter: btnFin, align: 'center', sortable: false },
            { label: 'Copiar', name: '', width: 15, formatter: btnConfig, align: 'center', sortable: false },
            { label: '', name: 'ItServicio', index: 'Servicio.IT', hidden: true },
            { label: '', name: 'HasInforme', hidden: true },
            { label: '', name: 'Aprobado', hidden: true },
            { label: '', name: 'AtrasadaInspeccion', hidden: true },
            { label: '', name: 'AtrasadaEntrega', hidden: true },
            { label: '', name: 'FechaEntrega', hidden: true },
            { label: '', name: 'Destinatario', hidden: true },
            { label: '', name: 'Fase1', hidden: true },
            { label: '', name: 'Califica', hidden: true },
            { label: '', name: 'HasNextFase', hidden: true },
            { label: '', name: 'Fec', hidden: true },
            { label: '', name: 'Fvc', hidden: true },
            { label: '', name: 'FromCotizacion', hidden: true }
            
        ],
        sortname: 'FechaCreacion',
        sortorder: 'desc',
        viewrecords: true,
        height: 380,
        rowNum: 30,
        pager: "#pager",
        caption: 'Inspecciones registradas en el sistema',
        hoverrows: false,
        rowattr: function(rd){
            if (rd.AtrasadaEntrega)
                return {
                    'class': 'red-row'
                }
            if (rd.AtrasadaInspeccion)
                return {
                    'class': 'yellow-row'
                }
        },
        // grouping
        grouping: true,
        groupingView: {
            groupField: ["ItServicio"],
            groupColumnShow: [false],
            groupText: ["<span class='jqgrid-grouping-text'><b> IT: {0}</b></span>"],
            groupOrder: ["asc"],
            groupSummary: [false],
            groupCollapse: false
                    
        },
        onCellSelect: function (rowid, iCol, cellcontent, e) {
            var rowData = $(this).jqGrid('getRowData', rowid);
            SELECTED_INSPECCION_ROW = rowData;
            SELECTED_INSPECCION = rowData.Id;
            console.log('icol', iCol);
            switch(iCol)
            {
                case 9:
                    openDialogEdit(rowData);
                    break;
                case 10:
                    if (rowData.Fase1 == '2')
                        return;
                    openDialogEspecificEdit(rowData);
                    break;
                case 11:
                    if (rowData.Fase1 == '2')
                        return;
                    openDialogNormas(rowData.Id, rowData.It);
                    break;
                case 12:
                    if (rowData.Fase1 == '1')
                    {
                        combobox($('#chl-f-titulo'), { 1: 'titulosInspeccion', inspeccionId: rowData.Id }, 'Seleccione un título');
                        openCheckListDialog();
                    }
                        
                    else if(rowData.Fase1 == '2')
                        openCheckListF2(rowData);
                    break;
                case 13:
                    if (rowData.Fase1 == '2')
                        getObservacionesTecnicasF2(rowData);
                    else
                        getObservacionesTecnicas(rowData);
                    break;
                case 14:
                    openCalificacion(rowData);
                    break;
                case 16:
                    var message, clase;
                    if(rowData.Aprobar == '')
                    {
                        alertify.error('Aún no se ha revisado')
                        break;
                    }
                    if (rowData.Aprobado == 'true')
                        break;
                    openAprobacion(rowData);
                    //openAprobacionDialog(rowData, message, clase, false);
                    break;
                case 17:

                    openAlertGetInforme(rowData);
                    break;
                case 18:
                    alertify.confirm('Certel', '¿Está seguro de finalizar esta inspección?. Esto significa que se borrará todo registro de esta.',
                        function () {
                            $.ajax({
                                url: 'handlers/Inspecciones.ashx',
                                method: 'POST',
                                data: {
                                    1: 'deleteInspeccion',
                                    id: rowData.Id
                                },
                                success: function (result) {
                                    if (result.done) {
                                        $('#formFiltros').submit();
                                    }
                                    else
                                        alertify.alert('Certel', result.message, function () { });
                                }
                            })
                        }, function () { }).set('labels', { ok: 'Finalizar', cancel: 'Cancelar' });
                    break;
                case 19:
                    $('#copy-dialog').dialog('open');
                    break;
                default: break;

            }
        }
    }).navGrid('#pager',
                {
                    edit: false, add: false, del: false, search: false,
                    refresh: true, view: false, position: "left", cloneToTop: false
                });

    function btnEstado(cellvalue, options, rowObject)
    {
        var color = '';
        switch (rowObject.EstadoId) {
            case 1: color = "rojo"; break;
            case 2: color = "amarillo"; break;
            case 3: color = "verde"; break;
        }
        return '<div class="btn-grid ' + color + '"><i class="fa fa-circle"></i></div>';
    }
    function btnRevisado(cellvalue, options, rowObject)
    {
        if (cellvalue == 0)
            return '';
        if(cellvalue == 2)
            return '<div class="btn-grid " onclick="revisar(' + rowObject.Id + ')"><i class="fa fa-check" style="color:gray"></i></div>';
        if(cellvalue == 1)
            return '<div class="btn-grid "><i class="fa fa-check" style="color:green"></i></div>';
        return '';
    }
    function btnAprobado(cellvalue, options, rowObject) {
        if (cellvalue == 0)
            return '';
        if (cellvalue == 2)
            return '<div class="btn-grid "><i class="fa fa-check" style="color:gray"></i></div>';
        if (cellvalue == 1)
            return '<div class="btn-grid "><i class="fa fa-check" style="color:green"></i></div>';
        return '';
    }
    function fase(cellvalue, options, rowObject) {
        var color = cellvalue == 1 ? '#000000' : '#2196F3';
        return '<b class="fase" style="color: ' + color + '">' + cellvalue + '</b>';
    }
    function btnEspecifico(cellvalue, options, rowObject) {
        if (rowObject.Fase1 == 2)
            return '';
        return '<div class="btn-grid "><i class="fa fa-pencil-square-o" style="color: #E65100"></i></div>';
    }
    function btnEditar(cellvalue, options, rowObject)
    {
        return '<div class="btn-grid"><i class="fa fa-pencil" style="color: #795548"></i></div>';
    }
    function btnFin(cellvalue, options, rowObject)
    {
        return '<div class="btn-grid rojo"><i class="fa fa-trash-o"></i></div>';
    }
    function btnCheckList() {
        return '<div class="btn-grid verde"><i class="fa fa-check-square-o"></i></div>';
    }
    function btnObsTec(cellvalue, options, rowObject) {
        if (rowObject.Fase1 == 2)
            return '<div class="btn-grid"><i class="fa fa-commenting"></i></div>';
        return '<div class="btn-grid azul"><i class="fa fa-commenting"></i></div>';
    }
    function btnCalifica(cellvalue, options, rowObject) {
        var mano = 'fa fa-question';
        var naranjo = '#FF9800';
        var verde = '#4CAF50';
        var rojo = '#F44336';
        var negro = '#444';
        var color;
        switch(rowObject.Califica)
        {
            case 0:
                mano = 'fa fa-thumbs-down';
                color = rojo;
                break;
            case 1:
                mano = 'fa fa-thumbs-up';
                color = verde;
                break;
            case 2:
                mano = 'fa fa-thumbs-o-up';
                color = naranjo;
                break;
            default: color = negro;
        }

        return '<div class="btn-grid"><i class="' + mano +'" style="color: ' + color +'"></i></div>';
    }
    function btnNormas(cellvalue, options, rowObject)
    {
        if (rowObject.Fase1 == 2)
            return '';
        return '<div class="btn-grid amarillo"><i class="fa fa-gavel"></i></div>';
    }
    function btnInforme(cellvalue, options, rowObject) {
        var color = '';
        if (rowObject.Aprobado)
            color = 'rojo'
        else
            color = 'gris'
        return '<div class="btn-grid ' + color +'"><i class="fa fa-file-pdf-o"></i></div>';
    }
    function btnConfig()
    {
        return '<div class="btn-grid gris"><i class="fa fa fa-clone"></i></div>';
    }
    $('#formFiltros').submit(function (e) {
        e.preventDefault();
        jQuery("#grid")
            .setGridParam({
                page: 1,
                postData: {
                    1: 'grid',
                    it: $('#f_it').val(),
                    desde: $('#f_desde').val(),
                    hasta: $('#f_hasta').val(),
                    calificacion: $('#f_calificacion').val(),
                    roles: localStorage.getItem('roles'),
                    user: localStorage.getItem('user')
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

    $('#chl-filters-form').submit(function () {
        getCheckList(SELECTED_INSPECCION_ROW);
        return false;
    }).find('select').change(function () {
        $(this).closest('form').submit();
    });
    // add inspeccion submit

    $('#add-inspeccion-form').submit(function (e) {
        e.preventDefault();
        if (!validateForm($(this))) {
            alertify.error('Complete los datos requeridos');
            return false;
        }

        $.ajax({
            url: 'handlers/Servicios.ashx',
            data: {
                1: 'addInspeccion',
                ubicacion: $('#ai-ubicacion').val(),
                fechaInstalacion: $('#ai-fecha-instalacion').val(),
                fechaInspeccion: $('#ai-fecha-inspeccion').val(),
                aparato: $('#ai-aparato').val(),
                funcionamiento: $('#ai-tipo-funcionamiento').val(),
                destino: $('#ai-destino').val(),
                permiso: $('#ai-permiso-edificacion').val(),
                recepcion: $('#ai-recepcion-municipal').val(),
                altura: $('#ai-altura').val() || 0,
                ingeniero: $('#ai-ingeniero').val(),
                servicio: 0,
                itServicio: $('#ai-it').val(),
                nombre: $('#ai-nombre').val(),
                edificio: $('#ai-edificio').val(),
                numero: $('#ai-numero').val(),
                fec: $('#ai-fec').val(),
                fvc: $('#ai-fvc').val()
            },
            success: function (result) {
                if (result.done) {
                    alertify.success(result.message);
                    $('#formFiltros').submit();
                    if (CLOSE_DIALOG)
                        $('#add-inspeccion-dialog').dialog('close');
                }
                else
                {
                    alertify.error(result.message);
                    if (result.code == 1)
                        $('#ai-it')
                            .addClass('no-validate')
                            .focus();
                }
                    
            }
        });
        return false;

    });
    
    // edit inspeccion submit
    $('#edit-inspeccion-form').submit(function (e) {
        e.preventDefault();
        if (!validateForm($(this))) {
            alertify.error('Complete los datos requeridos');
            return false;
        }

        $.ajax({
            url: 'handlers/Servicios.ashx',
            data: {
                1: 'editInspeccion',
                ubicacion: $('#ei-ubicacion').val(),
                fechaInstalacion: $('#ei-fecha-instalacion').val(),
                fechaInspeccion: $('#ei-fecha-inspeccion').val(),
                aparato: $('#ei-aparato').val(),
                funcionamiento: $('#ei-tipo-funcionamiento').val(),
                destino: $('#ei-destino').val(),
                permiso: $('#ei-permiso-edificacion').val(),
                recepcion: $('#ei-recepcion-municipal').val(),
                altura: $('#ei-altura').val(),
                ingeniero: $('#ei-ingeniero').val(),
                nombre: $('#ei-nombre').val(),
                edificio: $('#ei-edificio').val(),
                numero: $('#ei-numero').val(),
                fec: $('#ei-fec').val(),
                fvc: $('#ei-fvc').val(),
                id: SELECTED_INSPECCION
            },
            success: function (result) {
                if (result.done) {
                    alertify.success(result.message);
                    $('#formFiltros').submit();
                    $('#edit-inspeccion-dialog').dialog('close');
                }
                else {
                    alertify.error(result.message);
                    
                }

            }
        });
        return false;

    });

    // edit ot submit
    $('#eot-form').submit(function (e) {
        e.preventDefault();
        if (!validateForm($(this))) {
            alertify.error('Complete los datos requeridos');
            return false;
        }
        $.ajax({
            url: 'handlers/Inspecciones.ashx',
            type: 'POST',
            data: { 1: 'editOT', id: SELECTED_OT.Id, text: $('#eot-observacion').val() },
            success: function(result)
            {
                if(result.done)
                {
                    $('#edit-observacion-tecnica-dialog').dialog('close');
                    getObservacionesTecnicas(SELECTED_INSPECCION_ROW);
                }
                else
                {
                    alertify.alert('Zientte', result.message, function () { });
                }
            }
        })
        return false;
    });
    // edit form aprobacion
    $('#add').click(function () {
        $('#add-inspeccion-dialog').dialog({
            modal: true, bgiframe: false, width: '80%', title: 'Nueva Inspección', draggable: true,
            resizable: false, closeOnEscape: true, autoOpen: true, show: 'clip', hide: 'puff',
            open: function (type, data) {
                // $(this).find('form')[0].reset();
            },
            close: function () {
                $('#add-inspeccion-form')[0].reset();
                $('.no-validate').removeClass('no-validate');
            },
            buttons: [
                {
                    id: 'btn-add-i2',
                    text: 'Guardar y Crear Nuevo',
                    click: function () {
                        $('#add-inspeccion-form').submit();
                        CLOSE_DIALOG = false;
                    }
                },
                {
                    id: 'btn-add-i',
                    text: 'Guardar y Cerrar',
                    click: function () {
                        $('#add-inspeccion-form').submit();
                        CLOSE_DIALOG = true;
                    }
                },
                {
                    id: 'btn-add-close',
                    text: 'Cancelar',
                    click: function () {
                        $(this).dialog('close');
                    }
                },
            ]
        });
    });
    $('#edit-inspeccion-specific-form').submit(function (e) {
        e.preventDefault();
        var serializeArray = $(this).serializeArray();
        var data = JSON.stringify(serializeArray);
        $.ajax({
            url: 'handlers/Inspecciones.ashx',
            method: 'post',
            data: {
                1: 'editSpecificDataInspeccion',
                info: data,
                inspeccion: SELECTED_INSPECCION,
                all: SPECIFIC_DATA_APPLY_ALL
            },
            success: function (result) {
                if (result.done) {
                    alertify.success(result.message);
                    $('#edit-inspeccion-specific-dialog').dialog('close');
                }
                else {
                    alertify.error(result.message);
                    
                }

            }
        });
        return false;

    });

    $('#copy-form').submit(function () {
        var values = [];
        var selected = $('[name="copy-inspection"]:checked').each(function() {
            values.push(this.value);
        });
        var copyToNew = $('[name="copy-new-inspection"]').bootstrapSwitch('state');
        console.log('copytonew', copyToNew);
        if (values.length == 0)
        {
            alertify.error('Seleccione algún ítem para copiar');
            return false;
        }
        if (!copyToNew && $('#copy-to-it-exist').val() == 0)
        {
            alertify.error('Seleccione un IT al cual desea copiar la inspección. De lo contrario seleccione la opción "Nueva"');
            return false;
        }

        alertify.confirm('Certel', '¿Está seguro que desea copiar la inspección?. No podrá deshacer esta acción. El proceso podría durar algunos minutos.',
                function () {
                    $.ajax({
                        url: 'handlers/Inspecciones.ashx',
                        type: 'POST',
                        data: { 1: 'copy', array: JSON.stringify(values), from: SELECTED_INSPECCION, toNew: copyToNew, to: $('#copy-to-it-exist').val() },
                        success: function (result) {
                            if (result.done) {
                                alertify.alert('Certel', result.message, function () { });
                                $('#copy-dialog').dialog('close');
                                $('#formFiltros').submit();
                            }
                        }
                    })
                },
                function () {

                }).set('labels', { ok: 'Copiar', cancel: 'Cancelar' });
        
        return false;
    });

    $('#add-photo').change(function () {
        setFoto(SELECTED_CARACTERISTICA, SELECTED_INSPECCION, this);
    });
    $('#add-photo2').change(function () {
        var self = this;
        readURL(this, 2);
        $('#add-photo-observacion-tecnica').dialog({
            modal: true, bgiframe: false, width: 'auto', title: 'Ingrese una observación', draggable: true,
            resizable: false, closeOnEscape: true, autoOpen: true, show: 'clip', hide: 'puff',
            position: { my: "center", at: "top", of: document, collision: "fit" },
            open: function () {

            },
            buttons: {
                "Guardar": function () {
                    if(!$('#add-observacion-tecnica').val())
                    {
                        alertify.error('Complete una observación');
                        return;
                    }
                    setFotoTecnica($('#add-observacion-tecnica').val(), SELECTED_INSPECCION, self);
                }
            }
        });
    });
    $('#see-photos-dialog').dialog({
        modal: true, bgiframe: false, width: '50%', title: 'Fotografías', draggable: true,
        resizable: false, closeOnEscape: false, autoOpen: false, show: 'clip', hide: 'puff',
        position: { my: "center", at: "top", of: document, collision: "fit" },
        open: function (type, data) {

        },
    });
    $('#save-obs-tec').click(function () {
        if(!$('#obs-tec').val())
        {
            alertify.error('Complete una observación');
            return;
        }
        saveObservacion(SELECTED_INSPECCION,  $('#obs-tec').val());
    });
    $('#observaciones-tecnicas-dialog').dialog({
        modal: true, bgiframe: false, width: 800, title: 'Observaciones Técnicas', draggable: true,
        resizable: false, closeOnEscape: true, autoOpen: false, show: 'clip', hide: 'puff',
        height: $(window).height(),
        close: function () {
            $('#obs-tec').val('');
        },
        buttons: [{
            id: 'closeobs',
            text: 'Cerrar',
            click: function () {
                $('#observaciones-tecnicas-dialog').dialog('close');
            }
        }]
    });
    $('#observaciones-tecnicas-dialog-f2').dialog({
        modal: true, bgiframe: false, width: 800, title: 'Observaciones Técnicas', draggable: true,
        resizable: false, closeOnEscape: true, autoOpen: false, show: 'clip', hide: 'puff',
        height: $(window).height(),
        close: function () {
            
        },
        buttons: [{
            id: 'closeobs',
            text: 'Cerrar',
            click: function () {
                $('#observaciones-tecnicas-dialog-f2').dialog('close');
            }
        }]
    });

    $('#copy-dialog').dialog({
        modal: true, bgiframe: false, width: 600, title: 'Copiar Inspección', draggable: true,
        resizable: false, closeOnEscape: true, autoOpen: false, show: 'clip', hide: 'puff',
        open: function(){
            combobox($('#copy-to-it-exist'), { 1: 'sameIt', id: SELECTED_INSPECCION }, 'Seleccione IT...')
            $('[name="copy-inspection"][value="5"]').bootstrapSwitch('disabled', SELECTED_INSPECCION_ROW.Fase1 == '2');
            $('[name="copy-inspection"][value="6"]').bootstrapSwitch('disabled', SELECTED_INSPECCION_ROW.Fase1 == '2');
        
        },
        close: function () {

        },
        buttons: [{
            id: 'copy',
            text: 'Copiar',
            click: function () {
                $('#copy-form').submit();
            }
        }]
    });
    // Resize Panels
    $(window).bind('resize', function (e) {

        // grid
        if (grid = $('.ui-jqgrid-btable:visible')) {
            grid.each(function (index) {
                gridId = $(this).attr('id');
                gridParentWidth = $('#gbox_' + gridId).parent().width();
                gridParentHeight = $('#gbox_' + gridId).parent().height();
                $('#' + gridId).setGridWidth(gridParentWidth);
                $('#' + gridId).setGridHeight(gridParentHeight - 100);
            });
        }

    }).trigger('resize');

});

function openAprobacion(row)
{
    alertify.confirm(
        'Certel',
        '¿Está seguro que desea aprobar este informe?',
        function () {
            $.ajax(
            {
                url: 'handlers/Inspecciones.ashx',
                type: 'POST',
                data: { 1: 'soloAprobar', id: row.Id, usuario: localStorage.getItem('user') },
                success: function (result) {
                    if (result.done)
                    {
                        alertify.alert('Certel', result.message, function () { });
                        $('#formFiltros').submit();
                    }
                    else
                        alertify.error(result.message);
                }
            });
        },
        function () {

        }).set('labels', { ok: 'APROBAR', cancel: 'CANCELAR' });
}
function openAlertGetInforme(rowData)
{
    $('#exists-informe-dialog').dialog({
            modal: true, bgiframe: false, width: '40%', title: 'Generar Informe', draggable: true,
            resizable: false, closeOnEscape: true, autoOpen: true, show: 'clip', hide: 'puff',
            close: function () {

            },
            buttons: {
                "Crear": function () {
                    /*if (rowData.AtrasadaEntrega == 'true' || rowData.FechaEntrega == '')
                    {*/
                        openAprobacionDialog(rowData, 'Indique una fecha de entrega', 'info', true);
                    /*}
                    else
                    {
                        createPdf(rowData);
                    }*/
                },
                //"Abrir el último informe creado": function () {
                //    startReport(rowData);
                //}
            }
        });
}
function openAprobacionDialog(row, message,clase, openAlert)
{
    $('#message')
        .removeClass()
        .addClass('alert alert-' + clase).text(message);
    console.log(row);
    $('#fecha_entrega').val(row.FechaEntrega);
    $('#destinatario').val(row.Destinatario);
    $('#aprobacion-dialog').dialog({
        modal: true, bgiframe: false, width: '40%', title: 'Generar Informe', draggable: true,
        resizable: false, closeOnEscape: true, autoOpen: true, show: 'clip', hide: 'puff',
        close: function () {

        },
        buttons: [
            {
                id: 'aprobar',
                text: 'Generar',
                click: function () {
                    if (!validateForm($('#aprobacion-form'))) {
                        alertify.error('Complete los datos requeridos');
                        return false;
                    }
                    aprobar(row, $('#fecha_entrega').val(), $('#destinatario').val(), openAlert);
                }
            }
        ]
    });
}
function readURL(input) {
    if (input.files && input.files[0]) {
        if (!input.files[0].type.match(/image.*/)) {
            alertify.error('Archivo no es una imagen');
            return;
        };
        var reader = new FileReader();
        var img;
        reader.onload = function (e) {
           $('#add-photo-observacion-tecnica-img').prop({ 'src': e.target.result, 'height': 400 });    
        };   
    }
    reader.readAsDataURL(input.files[0]);     
}
function aprobar(row, fecha, destinatario, openAlert) {
    $.ajax({
        url: 'handlers/SetInforme.ashx',
        method: 'post',
        data: { 1: 'aprobar', id: row.Id, fecha: fecha, destinatario: destinatario },
        success: function (result) {
            if (result.done) {
                alertify.alert('Certel', result.message, function () { });
                $('#aprobacion-dialog').dialog('close');
                $('#formFiltros').submit();
                createPdf(row);
            }
            else
                alertify.error(result.message)
        }
    });
}
function openCalificacion(row) {
    
    $('#calificacion-dialog').dialog({
        modal: true, bgiframe: false, width: '50%', title: '¿Califica esta Inspección?', draggable: true,
        resizable: false, closeOnEscape: true, autoOpen: true, show: 'clip', hide: 'puff',
        open: function () {
            $('input[name=calificacion]')
                .closest('label')
                .removeClass('active');
            var cal = row.Califica;
            var $radios = $('input:radio[name=calificacion]');
            console.log('value on open', cal);
            
            $radios
                .filter('[value="' + cal +'"]')
                .prop('checked', true)
                .closest('label')
                .addClass('active');
            $('#fase2').hide();
            $('#plazo').hide();
            if (cal == '2' || cal == '0') {
                $('#plazo').show();
                if (row.HasNextFase == 'false') {
                    $('#fase2').show();
                }
            }
            
            
            
        },
        close: function () {
            
            $('input:radio[name="calificacion"]').prop('checked', false);
        },
        buttons: [
            {
                id: 'savecd',
                text: 'Guardar',
                click: function () {
                    var value = $('input[name="calificacion"]:checked').val();
                    console.log('guardando calif', value);
                    if(parseInt(value) > -1 && parseInt(value) < 3)
                        saveCalificacion(row.Id, value);
                    else
                        alertify.error('Ingrese una calificación válida')
                }
            }
        ]
    });
}
function openDialogEdit(row) {
    $.ajax({
        url: 'handlers/Inspecciones.ashx',
        method: 'POST',
        data: {
            1: 'getInspeccion',
            id: row.Id
        },
        success: function (result) {
            if (result.done) {
                var rowData = result.data;
                $('#ei-it').val(rowData.ItServicio);
                $('#ei-ubicacion').val(rowData.Ubicacion);
                $('#ei-fecha-instalacion').val(rowData.FechaInstalacion);
                $('#ei-fecha-inspeccion').val(rowData.FechaInspeccion);
                $('#ei-aparato').val(rowData.Aparato);
                $('#ei-tipo-funcionamiento').val(rowData.TipoFuncionamiento);
                $('#ei-destino').val(rowData.DestinoId);
                $('#ei-permiso-edificacion').val(rowData.PermisoEdificacion);
                $('#ei-recepcion-municipal').val(rowData.RecepcionMunicipal);
                $('#ei-altura').val(rowData.Altura);
                $('#ei-ingeniero').val(rowData.Ingeniero);
                $('#ei-nombre').val(rowData.Nombre);
                $('#ei-numero').val(rowData.Numero);
                $('#ei-edificio').val(rowData.Edificio);
                $('#ei-fec').val(rowData.Fec);
                $('#ei-fvc').val(rowData.Fvc);
                $('#edit-inspeccion-dialog').dialog({
                    modal: true, bgiframe: false, width: '80%', title: 'Editar Inspección IT: ' + row.It, draggable: true,
                    resizable: false, closeOnEscape: true, autoOpen: true, show: 'clip', hide: 'puff',
                    open: function (type, data) {
                        // $(this).find('form')[0].reset();
                    },
                    close: function () {
                        $('#edit-inspeccion-form')[0].reset();
                        $('.no-validate').removeClass('no-validate');
                    },
                    buttons: [

                        {
                            id: 'btn-edit-i',
                            text: 'Guardar',
                            click: function () {
                                $('#edit-inspeccion-form').submit();

                            }
                        },
                        {
                            id: 'btn-edit-close',
                            text: 'Cancelar',
                            click: function () {
                                $(this).dialog('close');
                            }
                        },
                    ]
                });
            }
            else {
                alertify.error(result.message);
            }
        }
    });
}
function openDialogEspecificEdit(row) {


    $.ajax({
        url: 'handlers/Inspecciones.ashx',
        method: 'POST',
        data: {
            1: 'getSpecificDataInspeccion',
            id: row.Id
        },
        success: function (result) {
            if (result.done) {
                $('#edit-inspeccion-specific-form').empty();
                $(result.data).each(function (i, item) {
                    $('#edit-inspeccion-specific-form')
                        .append($('<div>')
                            .addClass('form-group col-xs-12 col-lg-4 col-md-6 col-sm-12')
                            .append($('<label>')
                                .text(item.Nombre))
                            .append($('<input>')
                                .prop('type', 'text')
                                .prop('maxlenght', 300)
                                .prop('name', item.Id)
                                .data('id', item.Id)
                                .val(item.Valor)
                                .addClass('form-control')));
                });

                $('#edit-inspeccion-specific-dialog').dialog({
                    modal: true, bgiframe: false, width: '80%', title: 'Editar Inspección IT: ' + row.It, draggable: true,
                    resizable: false, closeOnEscape: true, autoOpen: true, show: 'clip', hide: 'puff',
                    close: function () {
                        SPECIFIC_DATA_APPLY_ALL = false;
                    },
                    buttons: [
                        {
                            id: 'save-specific',
                            text: 'Guardar',
                            click: function () {
                                $('#edit-inspeccion-specific-form').submit();
                            }
                        },
                        {
                            id: 'save-specific-all',
                            text: 'Guardar y Aplicar a Todos',
                            click: function () {
                                SPECIFIC_DATA_APPLY_ALL = true;
                                $('#edit-inspeccion-specific-form').submit();
                            }
                        }

                    ]
                });
            }

        }
    })
}
function openDialogNormas(rowid, rowit) {
    $.ajax({
        url: 'handlers/Inspecciones.ashx',
        method: 'POST',
        data: {
            1: 'getNormas',
            id: rowid
        },
        success: function (result) {
            console.log('openDialogNormas', result);
            if (result.done) {
                $('#asign-normas-form').empty();
                $('#asign-informe-form').empty();
                $(result.normas).each(function (i, item) {
                    var check;
                    $('#asign-normas-form')
                        .append($('<div>')
                        .addClass('form-group form-group-normas')
                        .append(check = $('<input>')
                            .prop('id', 'chn_' + item.Id)
                            .prop('type', 'checkbox')
                            .prop('name', 'chn_' + item.Id))
                        .append($('<label>')
                            .addClass('norma')
                            .prop('for', 'chn_' + item.Id)
                            .text(item.Nombre)))
                        .submit(function (e) {
                            e.preventDefault();
                            return false;
                        });

                    $('[name="chn_' + item.Id + '"]').bootstrapSwitch({
                        onText: 'SÍ',
                        offText: 'NO',
                        state: item.Checked,
                        onSwitchChange: function (event, state) {
                            addOrRemoveNorma(rowid, item.Id, state);
                            $('[name="chn_' + item.Madre + '"]').bootstrapSwitch('state', state);
                            var hijas = item.Hijas;
                            $(hijas).each(function (h, hija) {
                                $('[name="chn_' + hija + '"]').bootstrapSwitch('state', state);
                            });
                        }
                    });

                });

                $(result.informes).each(function (i, item) {
                    var check;
                    $('#asign-informe-form')
                        .append($('<div>')
                        .addClass('form-group form-group-normas')
                        .append(check = $('<input>')
                            .prop('id', 'chi_' + item.Id)
                            .prop('type', 'radio')
                            .prop('name', 'informe')
                            .prop('checked', item.Cheched))
                        .append($('<label>')
                            .addClass('norma')
                            .prop('for', 'chi_' + item.Id)
                            .text(item.Nombre)))
                        .submit(function (e) {
                            e.preventDefault();
                            return false;
                        });

                    $('[name="informe"]').bootstrapSwitch({
                        onText: 'SÍ',
                        offText: 'NO',
                        state: item.Checked,
                        onSwitchChange: function (event, state) {
                            setTipoInforme(rowid, item.Id);
                        }
                    });

                });

                $('#asign-normas-dialog').dialog({
                    modal: true, bgiframe: false, width: 800, title: 'Asignar Normas Inspección IT: ' + rowit, draggable: true,
                    resizable: false, closeOnEscape: true, autoOpen: true, show: 'clip', hide: 'puff',
                    buttons: [
                        {
                            id: 'close-normas',
                            text: 'Cerrar',
                            click: function () {
                                $('#asign-normas-dialog').dialog('close');
                            }
                        }
                    ]
                });
            }

        }
    })



}
function revisar(id) {
    alertify.confirm('Certel',
                    '¿Está seguro que quiere dar este informe por REVISADO?'
                    , function () {
                        $.ajax({
                            url: 'handlers/SetInforme.ashx',
                            method: 'post',
                            data: { 1: 'revisar', id: id, user: localStorage.getItem('user') },
                            success: function (result) {
                                if (result.done) {
                                    alertify.alert('Certel', result.message, function () { });
                                    $('#formFiltros').submit();
                                    
                                }
                                else
                                    alertify.error(result.message)
                            }
                        })
                    },
                    function () {

                    }).set('labels', { ok: 'REVISADO', cancel: 'CANCELAR' });
}
function addOrRemoveNorma(inspeccionId, id, state)
{
    $.ajax({
        url: 'handlers/Inspecciones.ashx',
        method: 'POST',
        data: {
            1: 'addOrRemoveNorma',
            id: id,
            state: state,
            inspeccion: inspeccionId
        },
        success: function (result) {
            console.log(result);
        }
    });
}
function setTipoInforme(inspeccion, informe)
{
    $.ajax({
        url: 'handlers/Inspecciones.ashx',
        method: 'POST',
        global:false,
        data: { 1: 'setTipoInforme', inspeccion: inspeccion, informe: informe },
        success: function (result) {
            console.log(result);
        }
    });
}
function openCheckListDialog()
{
    $('#check-list-dialog').dialog({
        modal: true, bgiframe: false, width: '60%', title: 'Check-list Inspección', draggable: true,
        resizable: false, closeOnEscape: true, autoOpen: true, show: 'clip', hide: 'puff',
        position: { my: "center", at: "top", of: document, collision: "fit" }, height: $(window).height(),
        open: function () { $('#check-list-panel').empty(); },
        buttons: [
            {
                id: 'close-chl',
                text: 'Cerrar',
                click: function () {
                    $(this).dialog('close');
                }
            }
        ]
    });
}
function getCheckList(row)
{
    
    $.ajax({
        url: 'handlers/Inspecciones.ashx',
        method: 'POST',
        data: {
            1: 'getCheckList',
            id: row.Id,
            title: $('#chl-f-titulo').val() || 0
        },
        success: function (result) {
            console.log('openchecklist', result);
            if(result.done)
            {
                $('#check-list-panel').empty();
                
                $(result.data).each(function (n, nor) {
                    
                    var norma;
                    $('#check-list-panel')
                        .append(norma = $('<div>')
                                .prop('id', 'norma_' + nor.Id)
                                .addClass('chl-norma')
                                .append($('<h3>')
                                    .text(nor.Text)));
                    $(nor.Titulos).each(function (t, titulo) {
                        var title;
                        $(norma)
                            .append(title = $('<div>')
                                .addClass('chl-titulo')
                                .append($('<h3>')
                                    .text(titulo.Text)));

                        $(titulo.Requisitos).each(function (r, req) {
                            var requisito;
                            $(title)
                                .append(requisito = $('<div>')
                                .addClass('chl-requisito')
                                .append($('<h4>')
                                    .text(req.Text))
                                .append($('<p>')
                                    .text(titulo.Text + ' - ' + nor.Text)));
                            $(req.Caracteristicas).each(function (c, car) {
                                var caracteristica;
                                var botonera;
                                var cumplimiento = car.Cumplimiento;
                                var evaluacion = cumplimiento == null ? 0 : cumplimiento.Id;
                                var observacion = cumplimiento == null ? '' : cumplimiento.Observacion;
                                var hasfotos = cumplimiento == null ? false : cumplimiento.HasFotos;
                                $(requisito)
                                    .append(caracteristica = $('<div>')
                                    .data('id', car.Id)
                                    .addClass('chl-caracteristica')
                                    .append($('<p>')
                                        .text(car.Text))
                                    .append(botonera = $('<div>')
                                        .addClass('chl-botonera')
                                        .append($('<div>')
                                            .addClass('chl-btn')
                                    // OK
                                            .append($('<button>')
                                                .data('caracteristica', car.Id)
                                                .data('active', false)
                                                .addClass('btn btn-default chl-ok')
                                                .addClass(evaluacion == 1 ? 'chl-active' : '')
                                                .prop('title', 'OK')
                                                .tooltip()
                                                .append($('<i>')
                                                    .addClass('fa fa-2x fa-check-circle'))
                                                .click(function () {
                                                    
                                                    setCumplimiento(1, car.Id, row.Id, this, botonera);
                                                })))
                                    // NA
                                        .append($('<div>')
                                            .addClass('chl-btn')
                                            .append($('<button>')
                                                .data('caracteristica', car.Id)
                                                .data('active', false)
                                                .addClass('btn btn-default chl-na')
                                                .addClass(evaluacion == 2 ? 'chl-active' : '')
                                                .prop('title', 'No Aplica')
                                                .tooltip()
                                                .append($('<i>')
                                                    .addClass('fa fa-2x fa-minus-circle'))
                                                .click(function () {
                                                    setCumplimiento(2, car.Id, row.Id, this, botonera);
                                                })))
                                    // NC
                                        .append($('<div>')
                                            .addClass('chl-btn')
                                            .append($('<button>')
                                                .data('caracteristica', car.Id)
                                                .data('active', false)
                                                .addClass('btn btn-default chl-nc')
                                                .addClass(evaluacion == 3 ? 'chl-active' : '')
                                                .prop('title', 'Defecto Grave')
                                                .tooltip()
                                                .append($('<i>')
                                                    .addClass('fa fa-2x fa-times-circle'))
                                                .click(function () {
                                                    setCumplimiento(3, car.Id, row.Id, this, botonera);
                                                })))
                                        // NC
                                        .append($('<div>')
                                            .addClass('chl-btn')
                                            .append($('<button>')
                                                .data('caracteristica', car.Id)
                                                .data('active', false)
                                                .addClass('btn btn-default chl-nc2')
                                                .addClass(evaluacion == 6 ? 'chl-active' : '')
                                                .prop('title', 'Defecto Leve')
                                                .tooltip()
                                                .append($('<i>')
                                                    .addClass('fa fa-2x fa-times-circle'))
                                                .click(function () {
                                                    setCumplimiento(3, car.Id, row.Id, this, botonera);
                                                })))

                                    // Observaciones
                                        .append($('<div>')
                                            .addClass('chl-btn')
                                            .append($('<button>')
                                                .data('caracteristica', car.Id)
                                                .data('active', false)
                                                .data('texto', observacion)
                                                .addClass('btn btn-default chl-obs')
                                                .addClass(observacion == '' || observacion == null ? '' : 'chl-active')
                                                .prop('title', 'Agregar Observación')
                                                .tooltip()
                                                .append($('<i>')
                                                    .addClass('fa fa-2x fa-commenting'))
                                                .click(function () {
                                                    SELECTED_CARACTERISTICA = car.Id;
                                                    setObs(car.Id, row.Id, $(this));
                                                })))

                                    //// Foto
                                    //    .append($('<div>')
                                    //        .addClass('fileUpload chl-btn')
                                    //        .append($('<div>')
                                    //            .data('caracteristica', car.Id)
                                    //            .data('active', false)
                                    //            .addClass('btn btn-default chl-foto')
                                    //            .prop('title', 'Agregar Fotografía')
                                    //            .tooltip()
                                    //            .append($('<i>')
                                    //                .addClass('fa fa-2x fa-camera'))
                                    //            .append($('<i>')
                                    //                .addClass('fa fa-plus'))
                                    //            .append($('<input>')
                                    //                .prop('type', 'file')
                                    //                .addClass('upload')
                                    //                .prop('accept', '.jpg, .gif, .png, .jpeg')
                                    //                .change(function(){
                                    //                    setFoto(car.Id, row.Id, this);
                                    //                }))
                                    //                .append($('<div>')
                                    //                    .addClass('pull-right')
                                    //                    .click(function () {
                                    //                        alert('ver fotos');
                                    //                    })
                                    //                    )
                                    //            ))
                                                

                                      // Imagenes
                                        .append($('<div>')
                                            .addClass('chl-btn')
                                            .append($('<button>')
                                                .data('caracteristica', car.Id)
                                                .data('active', hasfotos)
                                                .addClass('btn btn-default chl-img')
                                                .addClass(hasfotos ? 'chl-active' : '')
                                                .prop('title', 'Ver fotografías')
                                                .tooltip()
                                                .append($('<i>')
                                                    .addClass('fa fa-2x fa-image'))
                                                .click(function () {
                                                    SELECTED_CARACTERISTICA = car.Id;
                                                    SELECTED_INSPECCION = row.Id;
                                                    seePhotos(car.Id, row.Id, $(this));
                                                })))
                                        ));

                                
                            });
                        });
                    });
                    
                });
                
            }
            else
            {
                if(result.code == 1)
                {
                    alertify.confirm('Certel S.A',
                                        result.message,
                                        function () {
                                            openDialogNormas(row.Id, row.It);
                                        },
                                        function () {

                                        })
                                        .set('labels', { ok: 'OK', cancel: 'CANCELAR' });
                }
            }
        }
    });
}
function openCheckListF2(row) {
   
    $.ajax({
        url: 'handlers/Inspecciones.ashx',
        method: 'POST',
        data: {
            1: 'getCheckListF2',
            id: row.Id
        },
        success: function (result) {
            console.log('openchecklist', result);
            if (result.done) {
                $('#check-list-panel-f2').empty();

                $(result.data).each(function (n, nor) {

                    var norma;
                    $('#check-list-panel-f2')
                        .append(norma = $('<div>')
                                .prop('id', 'norma_' + nor.Id)
                                .addClass('chl-norma')
                                .append($('<h3>')
                                    .text(nor.Text)));
                    $(nor.Titulos).each(function (t, titulo) {
                        var title;
                        $(norma)
                            .append(title = $('<div>')
                                .addClass('chl-titulo')
                                .append($('<h3>')
                                    .text(titulo.Text)));

                        $(titulo.Requisitos).each(function (r, req) {
                            var requisito;
                            $(title)
                                .append(requisito = $('<div>')
                                .addClass('chl-requisito')
                                .append($('<h4>')
                                    .text(req.Text))
                                .append($('<p>')
                                    .text(titulo.Text + ' - ' + nor.Text)));
                            $(req.Caracteristicas).each(function (c, car) {
                                var caracteristica;
                                var botonera;
                                var cumplimiento = car.Cumplimiento;
                                var evaluacion = cumplimiento == null ? 0 : cumplimiento.Id;
                                var observacion = cumplimiento == null ? '' : cumplimiento.Observacion;
                                $(requisito)
                                    .append(caracteristica = $('<div>')
                                    .data('id', car.Id)
                                    .addClass('chl-caracteristica')
                                    .append($('<p>')
                                        .text(car.Text))
                                    .append(botonera = $('<div>')
                                        .addClass('chl-botonera')
                                        .append($('<div>')
                                            .addClass('chl-btn')
                                    // OK
                                            .append($('<button>')
                                                .data('caracteristica', car.Id)
                                                .data('active', false)
                                                .addClass('btn btn-default chl-ok')
                                                .addClass(evaluacion == 4 ? 'chl-active' : '')
                                                .prop('title', 'CORREGIDO')
                                                .tooltip()
                                                .append($('<i>')
                                                    .addClass('fa fa-2x fa-check-circle'))
                                                .click(function () {

                                                    setCumplimiento(4, car.Id, row.Id, this, botonera);
                                                })))
                                    
                                    // NC
                                        .append($('<div>')
                                            .addClass('chl-btn')
                                            .append($('<button>')
                                                .data('caracteristica', car.Id)
                                                .data('active', false)
                                                .addClass('btn btn-default chl-nc')
                                                .addClass(evaluacion == 5 ? 'chl-active' : '')
                                                .prop('title', 'NO CORREGIDO')
                                                .tooltip()
                                                .append($('<i>')
                                                    .addClass('fa fa-2x fa-times-circle'))
                                                .click(function () {
                                                    setCumplimiento(5, car.Id, row.Id, this, botonera);
                                                })))
                                      
                                        ));
                            });
                        });
                    });

                });




                $('#check-list-dialog-f2').dialog({
                    modal: true, bgiframe: false, width: '60%', title: 'Check-list Inspección', draggable: true,
                    resizable: false, closeOnEscape: true, autoOpen: true, show: 'clip', hide: 'puff',
                    position: { my: "center", at: "top", of: document, collision: "fit" }, height: $(window).height(),
                    open: function () {


                    },
                    buttons: [
                        {
                            id: 'close-chl',
                            text: 'Cerrar',
                            click: function () {
                                $(this).dialog('close');
                            }
                        }
                    ]
                });


            }
            else {
                if (result.code == 1) {
                    alertify.confirm('Certel S.A',
                                        result.message,
                                        function () {
                                            openDialogNormas(row.Id, row.It);
                                        },
                                        function () {

                                        })
                                        .set('labels', { ok: 'OK', cancel: 'CANCELAR' });
                }
            }
        }
    });
}
function getObservacionesTecnicas(row) {
    $.ajax({
        url: 'handlers/Inspecciones.ashx',
        data: {
            1: 'getObservacionesTecnicas',
            inspeccionId: row.Id,
        },
        success: function(result)
        {
            $('#obs-tec-list').empty();
            $(result).each(function (i, item) {
                var li, img;
                $('#obs-tec-list')
                    .append(li = $('<li>')
                        .addClass('list-group-item')
                        .append($('<div>')
                        .append($('<a>')
                            .prop('href', item.Image)
                            .append(img = $('<img>')
                                .addClass('img-rounded')
                                .prop('height', 100)
                                .prop('src', item.Image)
                                .error(function () {
                                    $(this).prop(
                                        {
                                            'src': 'fotos/default.jpg',
                                            'height': 100
                                        });
                                })))
                            .magnificPopup({
                                type: 'image',
                                delegate: 'a',
                                gallery: {
                                    enabled: true
                                },
                            })
                            .append($('<span>')
                                .addClass('spanObservacion')
                                .text(item.Texto)))
                                
                        .append($('<span>')
                            .addClass('btnDelete')
                                .append($('<i>')
                                .addClass('fa fa-remove')
                                .click(function () {
                                    removeObservacionTecnica(item.Id, li);
                                })))
                        .append($('<span>')
                            .addClass('btnEdit')
                                .append($('<i>')
                                .addClass('fa fa-pencil')
                                .click(function () {
                                    SELECTED_OT = item;
                                    editarObservacionTecnica();
                                })))
                        );

                    //.append($('<a>').
                    //            .prop('href', item.URL)
                    //        .append($('<img>')
                    //            .addClass('img-rounded img-responsive img-thumbnail')
                    //            .prop('src', item.URL)
                    //            .prop('width', 180)
                    //            .error(function () {
                    //                this.src = 'fotos/default.jpg'
                    //            })))
                    //            .magnificPopup({
                    //                type: 'image',
                    //                delegate: 'a',
                    //                gallery: {
                    //                    enabled: true
                    //                },
                    //            })

                if (item.Image == '')
                    img.hide();
                //);
            });
            if(!$('#observaciones-tecnicas-dialog').dialog('isOpen'))
                $('#observaciones-tecnicas-dialog').dialog('open');
        }
    })
    
}
function getObservacionesTecnicasF2(row) {
    $.ajax({
        url: 'handlers/Inspecciones.ashx',
        data: {
            1: 'getObservacionesTecnicasF2',
            inspeccionId: row.Id,
        },
        success: function (result) {
            $('#obs-tec-list-f2').empty();
            $(result).each(function (i, item) {
                var li, img,div;
                $('#obs-tec-list-f2')
                    .append(li = $('<li>')
                        .addClass('list-group-item')
                        .append($('<div>')
                            .append($('<a>', { href: item.Image })                               
                                .append(img = $('<img>', {
                                    class: 'img-rounded',
                                    height: 100,
                                    src: item.Image,
                                    error: function () {
                                        $(this).prop(
                                        {
                                            src: 'fotos/default.jpg',
                                            height: 100
                                        });
                                    }
                                })))
                                .magnificPopup({
                                    type: 'image',
                                    delegate: 'a',
                                    gallery: {
                                        enabled: true
                                    },
                                })
                            .append($('<span>')
                                .addClass('spanObservacionf2')
                                .text(item.Texto))
                            .append(div = $('<div>')
                                .addClass('btnCorregido')
                                .addClass(item.Corregido ? 'corregido' : '')
                                    .append($('<i>')
                                    .addClass('fa fa-check')
                                    )
                                    .click(function () {
                                       
                                        corregirOtF2(item.Id, div, $(this).hasClass('corregido'));
                                    }))));

                if (item.Image == '')
                    img.hide();
                //);
            });
            if (!$('#observaciones-tecnicas-dialog-f2').dialog('isOpen'))
                $('#observaciones-tecnicas-dialog-f2').dialog('open');
        }
    })

}
function corregirOtF2(id, div, ok)
{
    $.ajax({
        url: 'handlers/Inspecciones.ashx',
        type: 'POST',
        data: { 1: 'corregirOtF2', id: id, ok: !ok },
        success: function (result) {
            if(result.done)
            {
                if (result.ok)
                    div.addClass('corregido');
                else
                    div.removeClass('corregido');
            }
        }
    })
}
function editarObservacionTecnica()
{
    
    $('#eot-observacion').val(SELECTED_OT.Texto);
    $('#edit-observacion-tecnica-dialog').dialog({
        modal: true, bgiframe: false, width: 800, title: 'Editar Observación Técnica', draggable: true,
        resizable: false, closeOnEscape: true, autoOpen: true, show: 'clip', hide: 'puff',
        close: function () {
            $('#eot-observacion').val('');
        },
        buttons: [
            {
                id: 'editot',
                text: 'Guardar',
                click: function () {
                    $('#eot-form').submit();
                }
            }
        ]
    })
}
function openDialogObservacionesTecnicas()
{
    
}
function removeObservacionTecnica(id, li)
{
    $.ajax({
        url: 'handlers/Inspecciones.ashx',
        method: 'POST',
        data: { 1: 'removeObservacionTecnica', id: id },
        success: function (result) {
            if (result.done) {
                alertify.success(result.message);
                li.remove();
            }
            else {
                alertify.error(result.message);
            }
        }
    });
}
function saveObservacion(id, texto)
{
    $.ajax({
        url: 'handlers/Inspecciones.ashx',
        method: 'POST',
        data: { 1: 'saveObservacionTecnica', inspeccionId: id, texto: texto },
        success: function (result) {
            if (result.done) {
                var row = { Id: id };
                getObservacionesTecnicas(row);
                $('#obs-tec').val('').focus();
            }
            else {
                alertify.error('Ha ocurrido un error');
            }

        }
    });
}
function saveCalificacion(id, val)
{
    if ((val == 0 || val == 2) && $('#diasplazo').val() == '')
    {
        alertify.error('Debe ingresar la cantidad de días de plazo que tiene el cliente para resolver observaciones');
        return;
    }
    $.ajax({
        url: 'handlers/Inspecciones.ashx',
        method: 'POST',
        data: { 1: 'saveCalificacion', inspeccionId: id, val: val, dias: $('#diasplazo').val(), creafase2: $('#crearfase2').bootstrapSwitch('state') },
        success: function(result)
        {
            if (result.done)
            {
                $('#calificacion-dialog').dialog('close');
                alertify.alert('Certel', result.message, function () {
                    $('#formFiltros').submit();
                });
                
            }
            else
            {
                alertify.error('Ha ocurrido un error');
            }
                
        }
    })
}
function setCumplimiento(cum, car, insp, btn, botonera)
{
    $(botonera)
        .find('button')
        .filter(function () {
            return !$(this).hasClass('chl-obs')
        })
        .each(function () {
            $(this).removeClass('chl-active');
        });
    
    $.ajax({
        url: 'handlers/Inspecciones.ashx',
        method: 'post',
        data: {
            1: 'setCumplimiento',
            cumplimiento: cum,
            caracteristica: car,
            inspeccion: insp
        },
        global: false,
        success: function(result)
        {
            if (result.done)
            {
                $(btn).addClass('chl-active');
            }

            console.log(result);
        }
    })
}
function setObs(car, insp, btn) {
    var obs = $(btn).data('texto');
    console.log(btn, obs);
    $('#wod-observacion').val(obs);
    $('#writing-observacion-dialog').dialog({
        modal: true, bgiframe: false, width: '50%', title: 'Escriba una observación', draggable: true,
        resizable: false, closeOnEscape: true, autoOpen: true, show: 'clip', hide: 'puff',
        open: function (type, data) {
             
        },
        close: function () {
            $(this).find('form')[0].reset();
           
        },
        buttons: [

            {
                id: 'btn-writing-o',
                text: 'Guardar',
                click: function () {
                    $.ajax({
                        url: 'handlers/Inspecciones.ashx',
                        data: {
                            1: 'writeObservacion',
                            observacion: $('#wod-observacion').val(),
                            inspeccion: insp,
                            caracteristica: car
                        },
                        success: function (result) {
                            console.log(result);
                            if (result.done) {
                                if ($('#wod-observacion').val())
                                    $(btn).addClass('chl-active');
                                $('#writing-observacion-dialog').dialog('close');
                                $(btn).data('texto', $('#wod-observacion').val());
                            }
                            else {
                                alertify.error(result.message);
                            }
                        }
                    })

                }
            },
            {
                id: 'btn-edit-close',
                text: 'Cancelar',
                click: function () {
                    $(this).dialog('close');
                }
            },
        ]
    });
}
function setFoto(car, insp, input) {
    var file = input.files[0];
    
    name = file.name;
    size = file.size;
    type = file.type;

    if (file.name.length < 1) {
    }
    else if (file.size > 7000000) {
        alertify.alert('Certel', 'El archivo es demasiado pesado. Intente con otro.', function () { });
    }
    else if (file.type != 'image/png' && file.type != 'image/jpg' && file.type != 'image/gif' && file.type != 'image/jpeg') {
        alertify.alert('Certel', 'El archivo no tiene el formato correcto', function () { });
    }
    else {
        var formData = new FormData();
        formData.append(1, 'uploadImage');
        formData.append('file', file);
        formData.append('inspeccion', insp);
        formData.append('caracteristica', car);
        $.ajax({
            url: 'handlers/Inspecciones.ashx',
            type: 'post',
            data: formData,
            cache: false,
            contentType: false,
            processData: false,
            success: function (result) {
                if (result.done) {
                    alertify.success(result.message);
                    seePhotos(car, insp, input)
                }

                else
                    alertify.error(result.message);
            }

        });
    }
}
function setFotoTecnica(obs, insp, input) {
    var file = input.files[0];

    name = file.name;
    size = file.size;
    type = file.type;

    if (file.name.length < 1) {
        alertify.error("No se ha seleccionado una imagen");
    }
    else if (file.size > 7000000) {
        alertify.alert('Certel', "El archivo es demasiado pesado. Intente con otro.", function () { });
    }
    else if (file.type != 'image/png' && file.type != 'image/jpg' && file.type != 'image/gif' && file.type != 'image/jpeg') {
        alertify.alert('Certel', "El archivo no tiene el formato correcto", function () { });
    }
    else {
        var formData = new FormData();
        formData.append(1, 'uploadImageTecnica');
        formData.append('file', file);
        formData.append('obs', obs);
        formData.append('inspeccion', insp);
        $.ajax({
            url: 'handlers/Inspecciones.ashx',
            type: 'post',
            data: formData,
            cache: false,
            contentType: false,
            processData: false,
            success: function (result) {
                if (result.done) {
                    alertify.success(result.message);
                    $('#add-photo-observacion-tecnica').dialog('close');
                    var row = { Id: insp };
                    getObservacionesTecnicas(row);
                }

                else
                    alertify.error(result.message);
            }

        });
    }
}
function seePhotos(caracteristica, inspeccion, btn)
{
    $.ajax({
        url: 'handlers/Inspecciones.ashx',
        data: {
            1: 'getPhotos',
            caracteristica: caracteristica,
            inspeccion: inspeccion
        },
        success: function (result) {
            $('#panel-fotos').empty();
            if (result.done)
            {

                
                $(result.photos).each(function (i, item) {
                    $('#panel-fotos')
                        .append($('<div>')
                            .addClass('pf-foto')
                            .append($('<button>')
                                .prop('title', 'Eliminar Fotografía')
                                .tooltip()
                                .prop('type', 'button')
                                .prop('aria-label', 'Close')
                                .addClass('close')
                                    .append($('<span>')
                                        .prop('aria-hidden', true)
                                        .append($('<i>')
                                            .addClass('fa fa-times')))
                                .click(function () {
                                    removePhoto(item.Id, btn);
                                }))
                        .append($('<a>')
                            .prop('href', item.URL)
                            .append($('<img>')
                                .addClass('img-rounded img-responsive img-thumbnail')
                                .prop('src', item.URL)
                                .prop('width', 400)
                                .error(function () {
                                    this.src = 'fotos/default.jpg'
                                })))
                                .magnificPopup({
                                    type: 'image',
                                    delegate: 'a',
                                    gallery: {
                                        enabled: true
                                    },
                                }));
                    
                    //<button type="button" class="close" aria-label="Close"><span aria-hidden="true">&times;</span></button>
                });
                if (!$('#see-photos-dialog').is(':visible'))
                    $('#see-photos-dialog').dialog('open');
                $('#add-photo').prop('disabled', true);
            }
            else
            {
                if (result.code == 1)
                {
                    $('#panel-fotos')
                        .append($('<div>')
                            .addClass('no-results')
                            .text('No hay fotografías para esta característica'));
                    if (!$('#see-photos-dialog').is(':visible'))
                        $('#see-photos-dialog').dialog('open');
                    $('#add-photo').prop('disabled', false);
                }
                else
                    alertify.error(result.message);
            }
            
        }
    });
}
function removePhoto(id, btn)
{
    $.ajax({
        url: 'handlers/Inspecciones.ashx',
        method: 'post',
        data: {
            1: 'removePhoto',
            id: id
        },
        success: function(result)
        {
            if (result.done)
            {
                seePhotos(SELECTED_CARACTERISTICA, SELECTED_INSPECCION, btn);
                $('#add-photo').prop('disabled', false);
            }
                
            else
                alertify.error(result.message);
        }
    })
}
function openInforme(rowData)
{
    $.ajax({
        url: 'handlers/SetInforme.ashx',
        data: {
            1: 'getPlantilla',
            id: rowData.Id
        },
        success: function(result)
        {
            console.log(result);
        }
    })
}
function createPdf(rowData) {
    $.ajax({
        url: 'handlers/SetInforme.ashx',
        data: {
            1: 'createPdf',
            id: rowData.Id
        },
        success: function (result) {
            console.log(result);
            alertify.alert('Certel', result.message, function () {
                if (result.done)
                    openPDF('../pdf/' + result.path);
                
                if ($('#exists-informe-dialog').is(':visible'))
                    $('#exists-informe-dialog').dialog('close')
            });
            
                
        }
    });
}
function openPDF(url) {
    window.open(url, '_blank', 'fullscreen=yes')
}
function startReport(rowData)
{
    $.ajax({
        url: 'handlers/SetInforme.ashx',
        data: {
            1: 'start',
            id: rowData.Id
        },
        success: function (result) {
            alertify.alert('Certel', result.message, function () {
                if (result.done)
                    openPDF('../pdf/' + result.url);
                else
                    alertify.error(result.message);
            });
        }
    });
}
function getStructInform(rowData)
{
    $.ajax({
        url: 'handlers/Informes.ashx',
        data: {
            1: 'getStruct',
            id: rowData.Id
        },
        async: false,
        success: function (result) {
            if(result.done)
            {
                $('#panel-wizard').empty();
                $(result.data).each(function (i, item) {
                    var title = item.Title;
                    var text = item.Text;
                    $('#panel-wizard')
                        .append($('<h3>')
                            .text(item.Nombre))
                        .append($('<section>')
                            .append($('<form>')
                                .append($('<div>')
                                    .addClass('form-group')
                                        .append($('<input>')
                                            .prop('type', 'text')
                                            .prop('id', 'title_' + item.Id)
                                            .addClass('form-control')
                                            .val(title)))
                                .append($('<div>')
                                    .addClass('form-group')
                                        .append($('<div>')
                                            .addClass('form-control')
                                            .val(text)))));

                    
                });
                $('#wizard-informe-dialog').dialog({
                    modal: true, bgiframe: false, width: '80%', maxHeight: $(window).height(), title: 'Nueva Inspección', draggable: true,
                    resizable: false, closeOnEscape: true, autoOpen: true, show: 'clip', hide: 'puff',
                    position: { my: "center", at: "top", of: document, collision: "fit" },
                    open: function (type, data) {
                        $('#title_0').val('JAJSDJASJD')
                        // $(this).find('form')[0].reset();
                    },
                    close: function () {

                    },
                })
                $('#panel-wizard').steps({
                    headerTag: "h3",
                    bodyTag: "section",
                    transitionEffect: "slideLeft",
                    autoFocus: true,

                    stepsOrientation: 'vertical',
                    labels: {
                        cancel: "Cancelar",
                        current: "current step:",
                        pagination: "Paginación",
                        finish: "Finalizar",
                        next: "Siguiente",
                        previous: "Anterior",
                        loading: "Cargando ..."
                    },
                    onStepChanging: function (event, currentIndex, newIndex) {
                        return true;
                    },
                    onStepChanged: function (event, currentIndex, priorIndex) {
                    },
                    onCanceled: function (event) { },
                    onFinishing: function (event, currentIndex) {
                        return true;
                    },
                    onFinished: function (event, currentIndex) { },
                    onInit: function () {
                    }
                });
            }
            else
            {

            }
        }
    });
}

