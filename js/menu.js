$(function () {
    var form, form2, div, div2;
    var menu = $('<nav>')
        .attr('id', '_Menu_container')
        .addClass('navbar')
        .addClass('navbar-default')
        .attr('role', 'navigation')

    menu
        .append($('<div>')
            .addClass('navbar-header')
            .append($('<button>')
                .addClass('navbar-toggle')
                .addClass('collapsed')
                .attr('data-toggle', 'collapse')
                .attr('data-target', '#bs-example-navbar-collapse-1')
                .attr('aria-expanded', false)
                .append($('<span>')
                    .addClass('sr-only')
                    .text('toggle navigation'))
                .append($('<span>')
                    .addClass('icon-bar'))
                .append($('<span>')
                    .addClass('icon-bar'))
                .append($('<span>')
                    .addClass('icon-bar')))
            .append($('<a>')
                .addClass('navbar-brand')
                .prop('href', 'Login.aspx')
                .append($('<img>')
                    .prop('src', 'css/images/logo3.png')
                    .prop('width', '40')
                )));
    var nav = $('<div>')
        .addClass('collapse')
        .addClass('navbar-collapse')
        .addClass('navbar-ex1-collapse')
        .attr('id', 'bs-example-navbar-collapse-1')
    var ul = $('<ul>')
        .addClass('nav')
        .addClass('navbar-nav');

    //var items = ["Normas", "Inspecciones", "Servicios", "Clientes", "Cotizaciones", "Usuarios"];
    var items = []
    var rolesStr = localStorage.getItem('roles')
    if (!rolesStr)
        return;
    var roles = [];
    roles = JSON.parse(rolesStr);
    var ingenieria = {
        name: "Ingeniería",
        items: [
            { name: "Inspecciones", url: "Inspecciones.aspx" },
            { name: "Servicios", url: "Servicios.aspx" },
            { name: "Normas", url: "Normas.aspx" },
            { name: "Calendario", url: "Calendar.aspx" },
        ]
    }
    var comercial = {
        name: "Comercial",
        items: [
            { name: "Cotizaciones", url: "Cotizaciones.aspx" },
            { name: "Clientes", url: "Clientes.aspx" },
            { name: "Calendario", url: "Calendar.aspx" }
        ]
    }
    items.push(ingenieria);
    if (roles.indexOf(4) > -1 || roles.indexOf(5) > -1) // comercial
        items.push(comercial)

    var admin = {
            name: 'Configuración',
            items: [
                { name: 'Usuarios', url: 'Usuarios.aspx' },
                { name: 'Listas de precio', url: 'ListasDePrecio.aspx' },
                { name: 'Costos logísticos', url: 'CostosLogisticos.aspx' },
                { name: 'Productos', url: 'Productos.aspx' }
            ]
        }
    if (roles.indexOf(5) > -1) {
        items.push(admin);
    }
    $.each(items, function(i, item) {
        var ulul
        ul.append($('<li>', { class: 'dropdown'})
            .append($('<a>', {
                class: 'menu-li-a dropdown-toggle',
                text: item.name,
                'data-toggle': 'dropdown'
            })
                .append($('<b>', { class: 'caret' })))
            .append(ulul = $('<ul>', { class: 'dropdown-menu' })))

        $.each(item.items, function (ii, iitem) {
            ulul.append($('<li>')
                    .append($('<a>', {text: iitem.name, href: iitem.url })))
        })

    })

    var userUl = $('<ul>')
        .addClass('nav')
        .addClass('navbar-nav')
        .addClass('navbar-right')
        .append($('<li>')
            .addClass('dropdown')
            .append($('<a>')
                .addClass('name-user')
                .addClass('dropdown-toggle')
                .attr('data-toggle', 'dropdown')
                .text(localStorage.getItem('name'))
                .append($('<b>')
                    .addClass('caret')))
            .append($('<ul>')
                .addClass('dropdown-menu')
                .append($('<li>')
                    .append($('<a>')
                        .text('Salir')
                        .prop('href', 'Login.aspx')))
                .append($('<li>')
                    .append($('<a>')
                        .text('Cambiar Contraseña')
                        .click(function () {
                            div = $('<div>')
                                .prop('id', 'changePass')
                                .append(form = $('<form>')
                                    .append($('<div>')
                                        .addClass('form-group')
                                        .append($('<label>')
                                            .prop('for', 'chp_pass'))
                                        .append($('<input>')
                                            .addClass('form-control')
                                            .prop('type', 'password')
                                            .prop('id', 'chp_pass')
                                            .prop('placeholder', 'Actual Contraseña')))
                                    .append($('<div>')
                                        .addClass('form-group')
                                        .append($('<label>')
                                            .prop('for', 'chp_pass1'))
                                        .append($('<input>')
                                            .addClass('form-control')
                                            .prop('type', 'password')
                                            .prop('id', 'chp_pass1')
                                            .prop('placeholder', 'Nueva Contraseña. (De 4 a 8 caracteres)')))
                                    .append($('<div>')
                                        .addClass('form-group')
                                        .append($('<label>')
                                            .prop('for', 'chp_pass2'))
                                        .append($('<input>')
                                            .addClass('form-control')
                                            .prop('type', 'password')
                                            .prop('id', 'chp_pass2')
                                            .prop('placeholder', 'Repita Nueva Contraseña. (De 4 a 8 caracteres)')))
                                    .submit(function () {
                                        console.log('submit');
                                        if (!$('#chp_pass').val() || !$('#chp_pass1').val() || !$('#chp_pass2').val()) {
                                            alertify.error('Complete los datos');
                                            return false;
                                        }
                                        if ($('#chp_pass1').val() != $('#chp_pass2').val()) {
                                            alertify.error('Las contraseñas no coindiden!');
                                            return false;
                                        }
                                        if ($('#chp_pass1').val().length < 4 || $('#chp_pass1').val().length > 8) {
                                            alertify.error('La contraseña debe tener entre 4 y 8 caracteres');
                                            return false;
                                        }

                                        $.ajax({
                                            url: 'handlers/ChangePass.ashx',
                                            method: 'post',
                                            data: {
                                                user: localStorage.getItem('user'),
                                                oldPass: $('#chp_pass').val(),
                                                newPass: $('#chp_pass1').val()
                                            },
                                            success: function (result) {

                                                if (result.code == 1) {
                                                    alertify.success(result.message);
                                                    div.dialog('close');
                                                }
                                                else {
                                                    alertify.error(result.message);
                                                }
                                            }
                                        })
                                        return false;
                                    }))
                                .appendTo($('body'))
                                .css('display', 'none')
                                .dialog({
                                    modal: true,
                                    title: 'Cambiar Contraseña',
                                    autoOpen: true,
                                    width: 400,
                                    close: function () {
                                        div.remove();
                                    },
                                    buttons: [
                                        {
                                            id: 'savenewpass',
                                            text: 'Cambiar',
                                            click: function () {
                                                form.submit();
                                            }
                                        }
                                    ]
                                });
                        })))
                .append($('<li>')
                    .append($('<a>')
                        .text('Cambiar Contraseña E-mail')
                        .click(function () {
                            div2 = $('<div>')
                                .prop('id', 'changePassEmail')
                                .append(form2 = $('<form>')
                                    .append($('<div>')
                                        .addClass('form-group')
                                        .append($('<label>')
                                            .prop('for', 'chp_pass1_mail'))
                                        .append($('<input>')
                                            .addClass('form-control')
                                            .prop('type', 'password')
                                            .prop('id', 'chp_pass1_mail')
                                            .prop('placeholder', 'Nueva Contraseña. (De 4 a 8 caracteres)')))
                                    .append($('<div>')
                                        .addClass('form-group')
                                        .append($('<label>')
                                            .prop('for', 'chp_pass2_mail'))
                                        .append($('<input>')
                                            .addClass('form-control')
                                            .prop('type', 'password')
                                            .prop('id', 'chp_pass2_mail')
                                            .prop('placeholder', 'Repita Nueva Contraseña. (De 4 a 8 caracteres)')))
                                    .submit(function () {
                                        console.log('submit');
                                        if (!$('#chp_pass1_mail').val() || !$('#chp_pass2_mail').val()) {
                                            alertify.error('Complete los datos');
                                            return false;
                                        }
                                        if ($('#chp_pass1_mail').val() != $('#chp_pass2_mail').val()) {
                                            alertify.error('Las contraseñas no coindiden!');
                                            return false;
                                        }
                                        

                                        $.ajax({
                                            url: 'handlers/ChangePassMail.ashx',
                                            method: 'post',
                                            data: {
                                                user: localStorage.getItem('user'),
                                                newPass: $('#chp_pass1_mail').val()
                                            },
                                            success: function (result) {

                                                if (result.code == 1) {
                                                    alertify.success(result.message);
                                                    div2.dialog('close');
                                                }
                                                else {
                                                    alertify.error(result.message);
                                                }
                                            }
                                        })
                                        return false;
                                    }))
                                .appendTo($('body'))
                                .css('display', 'none')
                                .dialog({
                                    modal: true,
                                    title: 'Cambiar Contraseña Email',
                                    autoOpen: true,
                                    width: 400,
                                    close: function () {
                                        div2.remove();
                                    },
                                    buttons: [
                                        {
                                            id: 'savenewpassmail',
                                            text: 'Cambiar',
                                            click: function () {
                                                form2.submit();
                                            }
                                        }
                                    ]
                                });
                        })))
            ))
    nav.append(ul);
    nav.append(userUl);
    menu.append(nav);
    $("#cssmenu")
        .append(menu)     
            
})
    //});

    



