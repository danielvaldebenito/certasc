var URL_HANDLER = 'handlers/Usuarios.ashx'
var SELECTED_USER = {}
$(document).ready(function () {

    $('#au-firma, #eu-firma').jqte();
    $("#grid").jqGrid({
        url: URL_HANDLER,
        postData: {
            1: 'grid',
            name: ''
        },
        mtype: "POST",
        autowidth: true,
        datatype: "json",
        shrinkToFix: true,
        ajaxGridOptions: { global: false },
        colModel: [
            { label: 'Nombre de usuario', name: 'Username', key: true, hidden: true },
            { label: 'Nombre', name: 'Nombre', index: 'IT', width: 25, align: 'center', sortable: false },
            { label: 'Apellido', name: 'Apellido', width: 30, align: 'center', sortable: true },
            { label: 'Cargo', name: 'Cargo', width: 35, align: 'center', sortable: true },
            { label: 'Email', name: 'Email', width: 35, align: 'center', sortable: true },
            { label: 'Fono', name: 'Fono', width: 35, align: 'center', sortable: true },
            { label: 'Celular', name: 'Celular', width: 35, align: 'center', sortable: true },
            { label: 'Firma', name: 'Firma', hidden: true },
            { label: 'Habilitado', name: 'Habilitado', width: 10, align: 'center', sortable: true, formatter: btnEnableDisable },
            { label: 'Editar', width: 10, align: 'center', sortable: true, formatter: btnEditar },
            { label: 'Roles', width: 10, align: 'center', sortable: true, formatter: btnRoles }
           
        ],
        sortname: 'Nombre',
        sortorder: 'asc',
        viewrecords: true,
        height: 350,
        rowNum: 30,
        pager: "#pager",
        caption: 'Usuarios registrados',
        onCellSelect: function (rowid, iCol, cellcontent, e) {
            SELECTED_USER = $(this).getRowData(rowid)
            if (iCol == 9) {
                openDialogEdit(SELECTED_USER)
            }
            if (iCol == 10) {

                openDialogRoles(SELECTED_USER)
            }
        }
    }).navGrid('#pager',
                {
                    edit: false, add: false, del: false, search: false,
                    refresh: true, view: false, position: "left", cloneToTop: false
                });


    // Buttons
    function btnRoles() {
        return '<div class="btn-grid"><i class="fa fa-cogs"></i></div>'
    }
    function btnEnableDisable(cellvalue, options, rowObject) {
        var color = cellvalue ? 'verde' : 'rojo'
        return '<div class="btn-grid ' + color + '" onclick="enableDisable(\'' + rowObject.Username + '\')"><i class="fa fa-circle"></i></div>'
        
    }
    function btnEditar(cellvalue, options, rowObject) {
        return '<div class="btn-grid"><i class="fa fa-pencil"></i></div>'
    }
    $('#add').click(function () {
        $('#add-dialog').dialog('open');
    })
    $('#add-rol').click(function () {
        if ($('#rol').val() == 0) {
            alertify.error('Seleccione un rol')
            return;
        }
        addRol(SELECTED_USER.Username, $('#rol').val())
    })
    // Dialogs

    $('#add-dialog').dialog({
        modal: true, bgiframe: false, width: '80%', title: 'Nuevo Usuario', draggable: true,
        resizable: false, closeOnEscape: true, autoOpen: false, show: 'clip', hide: 'puff',
        open: function (type, data) {
            // $(this).find('form')[0].reset();
        },
        close: function () {
            $('#add-form')[0].reset();
            $('.no-validate').removeClass('no-validate');
        },
        buttons: [
            {
                id: 'btn-add',
                text: 'Guardar',
                click: function () {
                    $('#add-form').submit();
                }
            },
            {
                id: 'btn-close',
                text: 'Cancelar',
                click: function () {
                    $(this).dialog('close');
                }
            },
        ]
    });
    $('#edit-dialog').dialog({
        modal: true, bgiframe: false, width: '80%', title: 'Editar Usuario', draggable: true,
        resizable: false, closeOnEscape: true, autoOpen: false, show: 'clip', hide: 'puff',
        open: function (type, data) {
            // $(this).find('form')[0].reset();
        },
        close: function () {
            $('#edit-form')[0].reset();
            $('.no-validate').removeClass('no-validate');
        },
        buttons: [
            {
                id: 'btn-edit',
                text: 'Guardar',
                click: function () {
                    $('#edit-form').submit();
                }
            },
            {
                id: 'btn-close',
                text: 'Cancelar',
                click: function () {
                    $(this).dialog('close');
                }
            },
        ]
    });
    $('#roles-dialog').dialog({
        modal: true, bgiframe: false, width: '80%', title: 'Roles Usuario', draggable: true,
        resizable: false, closeOnEscape: true, autoOpen: false, show: 'clip', hide: 'puff'
    });
    $('#add-form').submit(function () {
        if (!validateForm($(this))) {
            alertify.error('Complete los datos requeridos');
            return false;
        }
        if ($('#au-password').val() != $('#au-password1').val()) {
            alertify.error('Las contraseñas no coinciden');
            return false;
        }
        if ($('#au-password-mail').val() != $('#au-password1-mail').val()) {
            alertify.error('Las contraseñas del correo no coinciden');
            return false;
        }
        add($('#au-username').val(), $('#au-nombre').val(), $('#au-apellido').val(), $('#au-password').val(), $('#au-cargo').val(), $('#au-email').val(), $('#au-firma').val(), $('#au-password-mail').val(), $('#au-fono').val(), $('#au-celular').val())
        return false;
    });
    $('#edit-form').submit(function () {
        if (!validateForm($(this))) {
            alertify.error('Complete los datos requeridos');
            return false;
        }
        edit($('#eu-username').val(), $('#eu-nombre').val(), $('#eu-apellido').val(), $('#eu-cargo').val(), $('#eu-email').val(), $('#eu-firma').val(), $('#eu-fono').val(), $('#eu-celular').val())
        return false;
    });
    $('#formFiltros').submit(function () {
        reloadGrid();
        return false;
    });
})
function openDialogEdit(data) {
    $('#eu-username').val(data.Username);
    $('#eu-nombre').val(data.Nombre);
    $('#eu-apellido').val(data.Apellido);
    $('#eu-cargo').val(data.Cargo);
    $('#eu-email').val(data.Email);
    $('#eu-fono').val(data.Fono);
    $('#eu-celular').val(data.Celular);
    $('#eu-firma').jqteVal(data.Firma);
    $('#edit-dialog').dialog('open')
}
function openDialogRoles(data) {
    $.ajax({
        url: URL_HANDLER,
        data: { 1: 'getUser', username: data.Username },
        success: function (result) {
            console.log(result)
            if (result.done) {
                var roles = result.user.Roles
                combobox($('#rol'), { 1: 'roles', username: result.user.Username }, 'Seleccione...');
                $('#roles-list').empty()
                $.each(roles, function (r, rol) {
                    $('#roles-list')
                        .append($('<li>', { class: 'list-group-item' })
                            .append($('<span>', { text: rol.Rol }))
                            .append($('<button>', { class: 'btn btn-danger' })
                                    .click(function () {
                                        removeRol(data.Username, rol.RolId)
                                    })
                                .append($('<i>', { class: 'fa fa-trash', title:'Quitar rol' }).tooltip())))
                })
                if (!$('#roles-dialog').is(':visible'))
                    $('#roles-dialog').dialog('open');
            } else {
                alertify.error(result.message)
            }
        }
    })
}
var add = function (username, name, surname, pass, cargo, email, firma, passMail, fono, celular) {
    $.ajax({
        url: URL_HANDLER,
        data: {
            1: 'add',
            username: username,
            name: name,
            surname: surname,
            pass: pass,
            cargo: cargo,
            email: email,
            firma: firma,
            passMail: passMail,
            fono: fono,
            celular: celular
        },
        success: function (result) {
            console.log(result)
            if (result.done) {
                alertify.success(result.message)
                reloadGrid();
                $('#add-dialog').dialog('close');
            } else {
                alertify.error(result.message)
            }
        }
    })
}
var edit = function (username, name, surname, cargo, email, firma, fono, celular) {
    $.ajax({
        url: URL_HANDLER,
        data: {
            1: 'edit',
            username: username,
            name: name,
            surname: surname,
            cargo: cargo,
            email: email,
            firma: firma,
            fono: fono,
            celular: celular
        },
        success: function (result) {
            console.log(result)
            if (result.done) {
                alertify.success(result.message)
                reloadGrid();
                $('#edit-dialog').dialog('close');
            } else {
                alertify.error(result.message)
            }
        }
    })
}
var enableDisable = function (username) {
    $.ajax({
        url: URL_HANDLER,
        data: {
            1: 'enableDisable',
            username: username
        },
        success: function (result) {
            console.log(result)
            if (result.done) {
                alertify.success(result.message)
                reloadGrid();
            } else {
                alertify.error(result.message)
            }
        }
    })
}
var addRol = function (username, rol) {
    $.ajax({
        url: URL_HANDLER,
        data: {
            1: 'addRol',
            username: username,
            rol: rol
        },
        success: function (result) {
            console.log(result)
            if (result.done) {
                alertify.success(result.message)
                openDialogRoles({ Username: username })
                combobox($('#rol'), { 1: 'roles', username: SELECTED_USER.Username }, 'Seleccione...');
            } else {
                alertify.error(result.message)
            }
        }
    })
}
var removeRol = function (username, rol) {
    $.ajax({
        url: URL_HANDLER,
        data: {
            1: 'removeRol',
            username: username,
            rol: rol
        },
        success: function (result) {
            console.log(result)
            if (result.done) {
                alertify.success(result.message)
                openDialogRoles({ Username: username })
            } else {
                alertify.error(result.message)
            }
        }
    })
}
var getUser = function (username) {
    $.ajax({
        url: URL_HANDLER,
        data: {
            1: 'getUser',
            username: username
        },
        success: function (result) {
            console.log(result)
            if (result.done) {
                alertify.success(result.message)
                reloadGrid();
            } else {
                alertify.error(result.message)
            }
        }
    })
}
var reloadGrid = function () {
    $("#grid").jqGrid('setGridParam', {
        url: URL_HANDLER,
        postData: {
            1: 'grid',
            name: $('#f_name').val()
        }
    }).trigger('reloadGrid');
}
