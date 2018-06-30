

$(function () {

    
    // Ajax Start and Ajax Stop
    $(document).ajaxStart(function () {
        $('.lock-screen').show();
    })
    .ajaxStop(function () {
        $('.lock-screen').hide();
    });

    // Resize Panels
    $(window).bind('resize', function (e) {
        
        var data = {
            ww : $(this).width(),
            hw : $(this).height()
        }

        
        //if (data.ww < 960) return;
        //$('body')
        //    .find("[data-height-full]")
        //    .each(function () {
        //        var top = $(this).offset().top;
        //        $(this)
        //            .height(data.hw - top - 60);
        //    });
        
    }).trigger('resize');
    
    

});

function combobox(cmb, data, defaultText, async)
{
    async = async == undefined ? true : async;
    $.ajax({
        url: 'Handlers/Combobox.ashx',
        data: data,
        async: async,
        success: function(result)
        {
            if (!result.done)
            {
                console.log('No se pudo cargar el combobox', result.message);
                return;
            }
            $(cmb).empty();
            if(defaultText != '')
                $(cmb)
                    .append($('<option>')
                    .val(0)
                    .text(defaultText));
            $(result.data).each(function (i, item) {
                $(cmb).append($('<option>')
                        .val(item.Value)
                        .text(item.Text));
            });
        }
    })
}

function validateForm(miForm) {
    var validate = true;
    var first = null;
    $(':input', miForm).each(function () {
        $(this).removeClass('no-validate');
        var type = this.type;
        var tag = this.tagName.toLowerCase();
        var required = this.required;

        if ((type == 'text' || type == 'password' || tag == 'textarea' || type == 'number') && required && this.value == '')
        {
            validate = false;
            $(this).addClass('no-validate');
            if (first == null)
                first = this;
        }
        else if (tag == 'select' && required && this.value == '0')
        {
            $(this).addClass('no-validate');
            validate = false;
            if (first == null)
                first = this;
        }
            
    });
    $(first).focus();
    console.log(first);
    return validate;
}



$.datepicker.regional['es'] = {
    closeText: 'Cerrar',
    prevText: '<Ant',
    nextText: 'Sig>',
    currentText: 'Hoy',
    monthNames: ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'],
    monthNamesShort: ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'],
    dayNames: ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'],
    dayNamesShort: ['Dom', 'Lun', 'Mar', 'Mié', 'Juv', 'Vie', 'Sáb'],
    dayNamesMin: ['Do', 'Lu', 'Ma', 'Mi', 'Ju', 'Vi', 'Sá'],
    weekHeader: 'Sm',
    dateFormat: 'dd-mm-yy',
    firstDay: 1,
    isRTL: false,
    showMonthAfterYear: false,
    yearSuffix: ''
};
$.datepicker.setDefaults($.datepicker.regional['es']);

function numberWithPoints(x) {
    var parts = x.toString().split(".");
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ".");
    return parts.join(".");
}

function getParameterByName(name, url) {
    if (!url) url = window.location.href;
    name = name.replace(/[\[\]]/g, "\\$&");
    var regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
        results = regex.exec(url);
    if (!results) return null;
    if (!results[2]) return '';
    return decodeURIComponent(results[2].replace(/\+/g, " "));
}